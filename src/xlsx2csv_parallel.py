#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlsx2csv with multiprocessing support

여러 시트를 병렬로 처리하여 성능을 향상시킨 버전
"""

import os
import sys
from multiprocessing import Pool, cpu_count
import csv
from pathlib import Path

# 원본 xlsx2csv 모듈
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx2csv import Xlsx2csv, XlsxException


def process_single_sheet(args):
    """
    단일 시트 처리 함수 (멀티프로세싱 워커)
    
    Args:
        args: (xlsxfile, sheet_info, output_dir, options) 튜플
    
    Returns:
        (sheet_name, success, error_message) 튜플
    """
    xlsxfile, sheet_info, output_dir, options = args
    
    sheet_index = sheet_info['index']
    sheet_name = sheet_info['name']
    
    try:
        # 각 프로세스에서 독립적으로 xlsx 파일 열기
        with Xlsx2csv(xlsxfile, **options) as xlsx2csv:
            # 출력 파일 경로 설정
            if sys.version_info[0] == 2:
                safe_sheet_name = sheet_name.encode('utf-8')
            else:
                safe_sheet_name = sheet_name
            
            output_file = os.path.join(output_dir, f"{safe_sheet_name}.csv")
            
            # 해당 시트만 변환
            xlsx2csv.convert(output_file, sheetid=sheet_index)
            
            return (sheet_name, True, None)
    
    except Exception as e:
        return (sheet_name, False, str(e))


class Xlsx2csvParallel:
    """
    멀티프로세싱을 지원하는 Xlsx2csv 래퍼 클래스
    
    여러 시트를 병렬로 처리하여 성능을 향상시킵니다.
    """
    
    def __init__(self, xlsxfile, num_processes=None, **options):
        """
        Args:
            xlsxfile: xlsx 파일 경로
            num_processes: 사용할 프로세스 수 (None이면 CPU 코어 수)
            **options: Xlsx2csv 옵션
        """
        self.xlsxfile = xlsxfile
        self.options = options
        self.num_processes = num_processes or cpu_count()
        
        # xlsx 파일 정보 미리 로드
        with Xlsx2csv(xlsxfile, **options) as xlsx2csv:
            self.sheets = xlsx2csv.workbook.sheets
            self.num_sheets = len(self.sheets)
    
    def convert_parallel(self, outdir, filter_sheets=None, verbose=True):
        """
        여러 시트를 병렬로 변환
        
        Args:
            outdir: 출력 디렉토리
            filter_sheets: 처리할 시트 인덱스 리스트 (None이면 모든 시트)
            verbose: 진행상황 출력 여부
        
        Returns:
            성공한 시트 목록
        """
        # 출력 디렉토리 생성
        if not os.path.exists(outdir):
            os.makedirs(outdir)
        
        # 필터링된 시트 목록
        sheets_to_process = []
        for sheet in self.sheets:
            if filter_sheets is None or sheet['index'] in filter_sheets:
                # 숨김 시트 필터링
                if self.options.get('exclude_hidden_sheets', False):
                    if sheet.get('state') in ('hidden', 'veryHidden'):
                        continue
                sheets_to_process.append(sheet)
        
        num_sheets_to_process = len(sheets_to_process)
        
        if verbose:
            print(f"병렬 처리 시작:")
            print(f"  파일: {self.xlsxfile}")
            print(f"  총 시트: {num_sheets_to_process}")
            print(f"  프로세스 수: {min(self.num_processes, num_sheets_to_process)}")
        
        # 단일 시트면 병렬화 불필요
        if num_sheets_to_process == 1:
            if verbose:
                print("  단일 시트 - 순차 처리")
            result = process_single_sheet((
                self.xlsxfile,
                sheets_to_process[0],
                outdir,
                self.options
            ))
            return [result[0]] if result[1] else []
        
        # 프로세스 풀 생성 및 병렬 처리
        # 시트 수보다 많은 프로세스는 불필요
        num_workers = min(self.num_processes, num_sheets_to_process)
        
        # 각 워커에 전달할 인자 준비
        args_list = [
            (self.xlsxfile, sheet, outdir, self.options)
            for sheet in sheets_to_process
        ]
        
        # 병렬 처리 실행
        successful_sheets = []
        failed_sheets = []
        
        with Pool(processes=num_workers) as pool:
            results = pool.map(process_single_sheet, args_list)
        
        # 결과 집계
        for sheet_name, success, error in results:
            if success:
                successful_sheets.append(sheet_name)
                if verbose:
                    print(f"  ✓ {sheet_name}")
            else:
                failed_sheets.append((sheet_name, error))
                if verbose:
                    print(f"  ✗ {sheet_name}: {error}")
        
        if verbose:
            print(f"\n완료: {len(successful_sheets)}/{num_sheets_to_process} 시트 성공")
            if failed_sheets:
                print(f"실패: {len(failed_sheets)} 시트")
                for name, error in failed_sheets:
                    print(f"  - {name}: {error}")
        
        return successful_sheets
    
    def convert_parallel_merged(self, outfile, delimiter="--------", verbose=True):
        """
        여러 시트를 병렬로 변환 후 하나의 파일로 병합
        
        Args:
            outfile: 출력 CSV 파일 경로
            delimiter: 시트 구분자
            verbose: 진행상황 출력 여부
        """
        import tempfile
        import shutil
        
        # 임시 디렉토리 생성
        temp_dir = tempfile.mkdtemp(prefix='xlsx2csv_parallel_')
        
        try:
            # 각 시트를 임시 디렉토리에 병렬 변환
            successful_sheets = self.convert_parallel(
                temp_dir,
                verbose=verbose
            )
            
            if not successful_sheets:
                raise XlsxException("변환된 시트가 없습니다")
            
            # 시트들을 하나의 파일로 병합
            if verbose:
                print(f"\n시트 병합 중: {outfile}")
            
            with open(outfile, 'w', encoding=self.options.get('outputencoding', 'utf-8'), newline='') as out_f:
                for i, sheet in enumerate(self.sheets):
                    sheet_name = sheet['name']
                    
                    if sheet_name not in successful_sheets:
                        continue
                    
                    # 시트 구분자 추가 (첫 시트 제외)
                    if i > 0 and delimiter:
                        out_f.write(delimiter + f" {sheet['index']} - {sheet_name}\n")
                    
                    # 시트 파일 읽어서 병합
                    sheet_file = os.path.join(temp_dir, f"{sheet_name}.csv")
                    if os.path.exists(sheet_file):
                        with open(sheet_file, 'r', encoding=self.options.get('outputencoding', 'utf-8')) as sheet_f:
                            shutil.copyfileobj(sheet_f, out_f)
            
            if verbose:
                print(f"✓ 병합 완료: {outfile}")
        
        finally:
            # 임시 디렉토리 정리
            shutil.rmtree(temp_dir, ignore_errors=True)


def convert_xlsx_parallel(xlsxfile, outdir, num_processes=None, **options):
    """
    편의 함수: xlsx 파일을 병렬로 변환
    
    Args:
        xlsxfile: 입력 xlsx 파일
        outdir: 출력 디렉토리
        num_processes: 프로세스 수
        **options: Xlsx2csv 옵션
    
    Returns:
        성공한 시트 목록
    """
    converter = Xlsx2csvParallel(xlsxfile, num_processes, **options)
    return converter.convert_parallel(outdir, verbose=True)


def main():
    """간단한 CLI"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="xlsx2csv with multiprocessing support"
    )
    parser.add_argument('input', help='입력 xlsx 파일')
    parser.add_argument('output', help='출력 디렉토리 또는 파일')
    parser.add_argument(
        '-p', '--processes',
        type=int,
        default=None,
        help=f'프로세스 수 (기본값: {cpu_count()})'
    )
    parser.add_argument(
        '-m', '--merged',
        action='store_true',
        help='하나의 파일로 병합'
    )
    parser.add_argument(
        '-d', '--delimiter',
        default='--------',
        help='시트 구분자 (병합 모드용)'
    )
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='진행상황 출력 안함'
    )
    
    args = parser.parse_args()
    
    options = {
        'delimiter': ',',
        'outputencoding': 'utf-8'
    }
    
    try:
        converter = Xlsx2csvParallel(
            args.input,
            num_processes=args.processes,
            **options
        )
        
        if args.merged:
            converter.convert_parallel_merged(
                args.output,
                delimiter=args.delimiter,
                verbose=not args.quiet
            )
        else:
            converter.convert_parallel(
                args.output,
                verbose=not args.quiet
            )
        
        print("\n✓ 변환 완료")
        return 0
    
    except Exception as e:
        print(f"\n✗ 오류: {str(e)}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
