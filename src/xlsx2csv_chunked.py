#!/usr/bin/env python3
"""
xlsx2csv 청크 기반 병렬 처리 버전

대용량 단일 시트를 여러 청크로 분할하여 병렬 처리
"""

import os
import sys
import zipfile
import xml.sax
import time
import tempfile
from multiprocessing import Pool, cpu_count
from xlsx2csv import Xlsx2csv, Sheet

class ChunkedSheetParser(xml.sax.ContentHandler):
    """특정 행 범위만 처리하는 SAX 파서"""
    
    def __init__(self, xlsx2csv_instance, start_row, end_row, include_header=True):
        """
        Args:
            xlsx2csv_instance: Xlsx2csv 인스턴스
            start_row: 시작 행 번호 (1-based)
            end_row: 종료 행 번호 (1-based, inclusive)
            include_header: 헤더 포함 여부
        """
        xml.sax.ContentHandler.__init__(self)
        self.xlsx2csv = xlsx2csv_instance
        self.start_row = start_row
        self.end_row = end_row
        self.include_header = include_header
        self.current_row_num = 0
        self.rows_data = []
        
        # 원본 Sheet의 속성 복사
        self.in_sheet = False
        self.in_row = False
        self.in_cell = False
        self.in_value = False
        self.current_cell = None
        self.current_row = []
        
    def startElement(self, name, attrs):
        if name == 'sheetData':
            self.in_sheet = True
        elif name == 'row' and self.in_sheet:
            self.in_row = True
            self.current_row = []
            self.current_row_num = int(attrs.get('r', '0'))
        elif name == 'c' and self.in_row:
            # 행 범위 체크
            if self.include_header and self.current_row_num == 1:
                # 헤더는 항상 포함
                self.in_cell = True
                self.current_cell = {
                    'r': attrs.get('r', ''),
                    't': attrs.get('t', ''),
                    's': attrs.get('s', ''),
                    'value': ''
                }
            elif self.start_row <= self.current_row_num <= self.end_row:
                self.in_cell = True
                self.current_cell = {
                    'r': attrs.get('r', ''),
                    't': attrs.get('t', ''),
                    's': attrs.get('s', ''),
                    'value': ''
                }
        elif name == 'v' and self.in_cell:
            self.in_value = True
        elif name == 'is' and self.in_cell:
            self.in_value = True
    
    def endElement(self, name):
        if name == 'sheetData':
            self.in_sheet = False
        elif name == 'row' and self.in_row:
            self.in_row = False
            # 행 범위 내의 데이터만 저장
            if self.include_header and self.current_row_num == 1:
                self.rows_data.append(self.current_row[:])
            elif self.start_row <= self.current_row_num <= self.end_row:
                self.rows_data.append(self.current_row[:])
        elif name == 'c' and self.in_cell:
            self.in_cell = False
            if self.current_cell:
                self.current_row.append(self.current_cell)
            self.current_cell = None
        elif name == 'v' and self.in_value:
            self.in_value = False
        elif name == 't' and self.in_value:
            self.in_value = False
    
    def characters(self, content):
        if self.in_value and self.current_cell is not None:
            self.current_cell['value'] += content


def get_sheet_dimensions(xlsx_file, sheet_index=1):
    """
    시트의 차원(행 수, 열 수) 파악
    
    Returns:
        (max_row, max_col) 튜플
    """
    with zipfile.ZipFile(xlsx_file) as zf:
        # sheet 파일 찾기
        sheet_file = f'xl/worksheets/sheet{sheet_index}.xml'
        
        with zf.open(sheet_file) as f:
            # dimension 태그 찾기
            content = f.read(10000).decode('utf-8')  # 처음 일부만 읽기
            
            # <dimension ref="A1:O100000"/> 형식 파싱
            if 'dimension' in content:
                import re
                match = re.search(r'dimension ref="[A-Z]+\d+:([A-Z]+)(\d+)"', content)
                if match:
                    col_letter = match.group(1)
                    max_row = int(match.group(2))
                    
                    # 열 문자를 숫자로 변환 (A=1, B=2, ..., AA=27)
                    max_col = 0
                    for char in col_letter:
                        max_col = max_col * 26 + (ord(char) - ord('A') + 1)
                    
                    return max_row, max_col
            
            # dimension이 없으면 전체 스캔 (느림)
            # 단순화를 위해 기본값 반환
            return None, None


def process_chunk(args):
    """
    청크를 처리하는 워커 함수
    
    Args:
        args: (xlsx_file, sheet_index, start_row, end_row, chunk_output_file, options, include_header)
    
    Returns:
        (chunk_id, success, error_msg)
    """
    xlsx_file, sheet_index, start_row, end_row, chunk_output_file, options, include_header = args
    
    try:
        # Xlsx2csv 인스턴스 생성
        xlsx2csv = Xlsx2csv(xlsx_file, **options)
        
        # ZIP 파일 열기
        with zipfile.ZipFile(xlsx_file) as zf:
            sheet_file = f'xl/worksheets/sheet{sheet_index}.xml'
            
            with zf.open(sheet_file) as sheet_filehandle:
                # 청크 파서로 처리
                parser = ChunkedSheetParser(xlsx2csv, start_row, end_row, include_header)
                xml.sax.parse(sheet_filehandle, parser)
                
                # CSV로 저장
                with open(chunk_output_file, 'w', encoding='utf-8', newline='') as out:
                    import csv
                    writer = csv.writer(out, delimiter=options.get('delimiter', ','))
                    
                    # 각 행 처리
                    for row_cells in parser.rows_data:
                        row_data = []
                        for cell in row_cells:
                            # 셀 값 변환 (shared_strings, 숫자 등)
                            value = cell['value']
                            cell_type = cell.get('t', '')
                            
                            if cell_type == 's':  # shared string
                                try:
                                    idx = int(value)
                                    if idx < len(xlsx2csv.shared_strings.strings):
                                        value = xlsx2csv.shared_strings.strings[idx]
                                except:
                                    pass
                            
                            row_data.append(value)
                        
                        writer.writerow(row_data)
        
        return (chunk_output_file, True, None)
    
    except Exception as e:
        return (chunk_output_file, False, str(e))


def merge_chunks(chunk_files, output_file, remove_chunks=True):
    """
    청크 파일들을 하나로 병합
    
    Args:
        chunk_files: 청크 파일 경로 리스트 (순서대로)
        output_file: 최종 출력 파일
        remove_chunks: 병합 후 청크 파일 삭제 여부
    """
    with open(output_file, 'w', encoding='utf-8', newline='') as out:
        for i, chunk_file in enumerate(chunk_files):
            with open(chunk_file, 'r', encoding='utf-8') as f:
                if i == 0:
                    # 첫 번째 청크는 모두 복사 (헤더 포함)
                    out.write(f.read())
                else:
                    # 나머지 청크는 헤더 스킵
                    lines = f.readlines()
                    if len(lines) > 1:  # 헤더 제외하고 복사
                        out.writelines(lines[1:])
            
            # 청크 파일 삭제
            if remove_chunks:
                try:
                    os.remove(chunk_file)
                except:
                    pass


class Xlsx2csvChunked:
    """청크 기반 병렬 처리를 위한 래퍼 클래스"""
    
    def __init__(self, xlsx_file, **kwargs):
        self.xlsx_file = xlsx_file
        self.options = kwargs
        
    def convert_chunked(self, output_file, sheet_index=1, chunk_size=50000, num_workers=None):
        """
        청크 기반 병렬 변환
        
        Args:
            output_file: 출력 CSV 파일
            sheet_index: 시트 인덱스 (1-based)
            chunk_size: 청크당 행 수
            num_workers: 워커 프로세스 수 (None이면 CPU 수)
        """
        if num_workers is None:
            num_workers = cpu_count()
        
        print(f"\n{'='*60}")
        print(f"청크 기반 병렬 변환 시작")
        print(f"파일: {self.xlsx_file}")
        print(f"청크 크기: {chunk_size:,}행")
        print(f"워커 수: {num_workers}")
        print(f"{'='*60}\n")
        
        # 1. 시트 차원 파악
        print("1. 시트 분석 중...")
        max_row, max_col = get_sheet_dimensions(self.xlsx_file, sheet_index)
        
        if max_row is None:
            print("   ⚠️  차원 정보 없음, 전체 스캔 필요")
            # 간단한 구현을 위해 기본 청크 사용
            max_row = 100000
        
        print(f"   - 총 행 수: {max_row:,}")
        print(f"   - 총 열 수: {max_col}")
        
        # 2. 청크 범위 계산
        print("\n2. 청크 분할 중...")
        ranges = []
        start = 2  # 1은 헤더
        while start <= max_row:
            end = min(start + chunk_size - 1, max_row)
            ranges.append((start, end))
            start = end + 1
        
        num_chunks = len(ranges)
        print(f"   - 청크 수: {num_chunks}")
        for i, (s, e) in enumerate(ranges):
            print(f"   - Chunk {i+1}: 행 {s:,} ~ {e:,} ({e-s+1:,}행)")
        
        # 3. 청크별 병렬 처리
        print("\n3. 청크 병렬 처리 중...")
        
        temp_dir = tempfile.gettempdir()
        chunk_files = []
        args_list = []
        
        for i, (start_row, end_row) in enumerate(ranges):
            chunk_file = os.path.join(temp_dir, f"chunk_{i:04d}.csv")
            chunk_files.append(chunk_file)
            
            # 첫 번째 청크만 헤더 포함
            include_header = (i == 0)
            
            args_list.append((
                self.xlsx_file,
                sheet_index,
                start_row,
                end_row,
                chunk_file,
                self.options,
                include_header
            ))
        
        start_time = time.time()
        
        with Pool(processes=num_workers) as pool:
            results = pool.map(process_chunk, args_list)
        
        process_time = time.time() - start_time
        
        # 결과 확인
        success_count = sum(1 for _, success, _ in results if success)
        print(f"\n   ✅ 처리 완료: {success_count}/{num_chunks} 청크")
        print(f"   ⏱️  처리 시간: {process_time:.3f}초")
        
        if success_count < num_chunks:
            print("\n   ⚠️  실패한 청크:")
            for chunk_file, success, error in results:
                if not success:
                    print(f"      - {chunk_file}: {error}")
        
        # 4. 청크 병합
        print("\n4. 청크 병합 중...")
        merge_start = time.time()
        merge_chunks(chunk_files, output_file, remove_chunks=True)
        merge_time = time.time() - merge_start
        
        print(f"   ✅ 병합 완료: {merge_time:.3f}초")
        
        total_time = time.time() - start_time
        print(f"\n{'='*60}")
        print(f"✅ 전체 완료!")
        print(f"총 시간: {total_time:.3f}초 (처리: {process_time:.3f}초 + 병합: {merge_time:.3f}초)")
        print(f"{'='*60}\n")
        
        return total_time


def main():
    """CLI 인터페이스"""
    import argparse
    
    parser = argparse.ArgumentParser(description='xlsx2csv 청크 기반 병렬 변환')
    parser.add_argument('input_file', help='입력 xlsx 파일')
    parser.add_argument('output_file', help='출력 CSV 파일')
    parser.add_argument('--sheet', type=int, default=1, help='시트 인덱스 (기본: 1)')
    parser.add_argument('--chunk-size', type=int, default=50000, help='청크 크기 (기본: 50000)')
    parser.add_argument('--workers', type=int, default=None, help='워커 수 (기본: CPU 수)')
    parser.add_argument('--delimiter', default=',', help='CSV 구분자 (기본: ,)')
    
    args = parser.parse_args()
    
    converter = Xlsx2csvChunked(
        args.input_file,
        delimiter=args.delimiter
    )
    
    converter.convert_chunked(
        args.output_file,
        sheet_index=args.sheet,
        chunk_size=args.chunk_size,
        num_workers=args.workers
    )


if __name__ == '__main__':
    main()
