#!/usr/bin/env python3
"""
xlsx2csv 하이브리드 적응형 최적화 시스템

파일 특성을 분석하여 자동으로 최적 전략 선택:
1. 순차 처리 (작은 파일)
2. 시트 병렬 처리 (다중 시트)
3. 청크 병렬 처리 (대용량 단일 시트)
4. 복합 전략 (다중 시트 + 대용량 시트)
"""

import os
import zipfile
import time
from multiprocessing import cpu_count
from xlsx2csv import Xlsx2csv
from xlsx2csv_parallel import Xlsx2csvParallel
from xlsx2csv_chunked import Xlsx2csvChunked, get_sheet_dimensions

class FileAnalyzer:
    """파일 특성 분석"""
    
    def __init__(self, xlsx_file):
        self.xlsx_file = xlsx_file
        self.file_size = os.path.getsize(xlsx_file)
        self.file_size_mb = self.file_size / (1024 * 1024)
        
    def analyze(self):
        """
        파일 분석하여 전략 결정에 필요한 정보 수집
        
        Returns:
            {
                'file_size_mb': float,
                'num_sheets': int,
                'sheets': [{'index': int, 'name': str, 'rows': int, 'cols': int}],
                'total_rows': int
            }
        """
        print(f"\n{'='*60}")
        print(f"📊 파일 분석 중...")
        print(f"파일: {os.path.basename(self.xlsx_file)}")
        print(f"크기: {self.file_size_mb:.2f} MB")
        print(f"{'='*60}\n")
        
        sheets_info = []
        total_rows = 0
        
        with zipfile.ZipFile(self.xlsx_file) as zf:
            # 시트 목록 확인
            sheet_index = 1
            while True:
                sheet_file = f'xl/worksheets/sheet{sheet_index}.xml'
                try:
                    # dimension 파악
                    max_row, max_col = get_sheet_dimensions(self.xlsx_file, sheet_index)
                    
                    if max_row is None:
                        # dimension이 없으면 추정
                        info = zf.getinfo(sheet_file)
                        # 파일 크기로 대략적 행 수 추정 (1행 ≈ 150 bytes)
                        max_row = info.file_size // 150
                        max_col = 10
                    
                    sheets_info.append({
                        'index': sheet_index,
                        'name': f'Sheet{sheet_index}',
                        'rows': max_row,
                        'cols': max_col
                    })
                    
                    total_rows += max_row
                    
                    print(f"  Sheet {sheet_index}: {max_row:,}행 × {max_col}열")
                    
                    sheet_index += 1
                    
                except KeyError:
                    # 더 이상 시트 없음
                    break
        
        num_sheets = len(sheets_info)
        
        print(f"\n총 {num_sheets}개 시트, {total_rows:,}행")
        print(f"{'='*60}\n")
        
        return {
            'file_size_mb': self.file_size_mb,
            'num_sheets': num_sheets,
            'sheets': sheets_info,
            'total_rows': total_rows
        }


class StrategySelector:
    """최적 전략 선택"""
    
    # 임계값 설정
    SMALL_FILE_THRESHOLD_MB = 1.0  # 1MB 미만은 순차 처리
    LARGE_SHEET_THRESHOLD_ROWS = 50000  # 50K 행 이상은 청크 처리
    MIN_SHEETS_FOR_PARALLEL = 2  # 2개 이상 시트면 병렬 고려
    
    def __init__(self, file_info):
        self.file_info = file_info
        
    def select_strategy(self):
        """
        최적 전략 선택
        
        Returns:
            {
                'strategy': str,  # 'sequential', 'sheet_parallel', 'chunk_parallel', 'hybrid'
                'reason': str,
                'config': dict
            }
        """
        file_size_mb = self.file_info['file_size_mb']
        num_sheets = self.file_info['num_sheets']
        sheets = self.file_info['sheets']
        
        print(f"{'='*60}")
        print(f"🎯 전략 선택 중...")
        print(f"{'='*60}\n")
        
        # 1. 작은 파일 → 순차 처리
        if file_size_mb < self.SMALL_FILE_THRESHOLD_MB:
            return {
                'strategy': 'sequential',
                'reason': f'작은 파일 ({file_size_mb:.2f}MB < {self.SMALL_FILE_THRESHOLD_MB}MB)',
                'config': {}
            }
        
        # 2. 시트 분석
        large_sheets = [s for s in sheets if s['rows'] >= self.LARGE_SHEET_THRESHOLD_ROWS]
        small_sheets = [s for s in sheets if s['rows'] < self.LARGE_SHEET_THRESHOLD_ROWS]
        
        # 3. 단일 시트
        if num_sheets == 1:
            sheet = sheets[0]
            if sheet['rows'] >= self.LARGE_SHEET_THRESHOLD_ROWS:
                # 대용량 단일 시트 → 청크 병렬
                chunk_size = 50000
                num_chunks = (sheet['rows'] + chunk_size - 1) // chunk_size
                num_workers = min(num_chunks, cpu_count())
                
                return {
                    'strategy': 'chunk_parallel',
                    'reason': f'대용량 단일 시트 ({sheet["rows"]:,}행)',
                    'config': {
                        'chunk_size': chunk_size,
                        'num_workers': num_workers
                    }
                }
            else:
                # 작은 단일 시트 → 순차
                return {
                    'strategy': 'sequential',
                    'reason': f'작은 단일 시트 ({sheet["rows"]:,}행)',
                    'config': {}
                }
        
        # 4. 다중 시트
        if num_sheets >= self.MIN_SHEETS_FOR_PARALLEL:
            # 대용량 시트가 있는지 확인
            if len(large_sheets) > 0:
                # 하이브리드: 시트 병렬 + 청크 병렬
                return {
                    'strategy': 'hybrid',
                    'reason': f'{num_sheets}개 시트 (대용량 {len(large_sheets)}개 포함)',
                    'config': {
                        'large_sheets': large_sheets,
                        'small_sheets': small_sheets,
                        'num_workers': min(num_sheets, cpu_count()),
                        'chunk_size': 50000
                    }
                }
            else:
                # 시트 병렬만
                num_workers = min(num_sheets, cpu_count())
                return {
                    'strategy': 'sheet_parallel',
                    'reason': f'{num_sheets}개 시트 (모두 중소형)',
                    'config': {
                        'num_workers': num_workers
                    }
                }
        
        # 기본값: 순차
        return {
            'strategy': 'sequential',
            'reason': '기본 전략',
            'config': {}
        }


class Xlsx2csvHybrid:
    """하이브리드 적응형 변환기"""
    
    def __init__(self, xlsx_file, **kwargs):
        self.xlsx_file = xlsx_file
        self.options = kwargs
        
    def convert_auto(self, output_dir=None):
        """
        자동 최적화 변환
        
        Args:
            output_dir: 출력 디렉토리 (None이면 파일명 기반 생성)
        
        Returns:
            총 처리 시간 (초)
        """
        start_time = time.time()
        
        print(f"\n{'#'*70}")
        print(f"# xlsx2csv 하이브리드 적응형 변환")
        print(f"# {os.path.basename(self.xlsx_file)}")
        print(f"{'#'*70}\n")
        
        # 1. 파일 분석
        analyzer = FileAnalyzer(self.xlsx_file)
        file_info = analyzer.analyze()
        
        # 2. 전략 선택
        selector = StrategySelector(file_info)
        strategy_info = selector.select_strategy()
        
        strategy = strategy_info['strategy']
        reason = strategy_info['reason']
        config = strategy_info['config']
        
        print(f"선택된 전략: {strategy}")
        print(f"이유: {reason}")
        print(f"설정: {config}")
        print(f"\n{'='*60}\n")
        
        # 3. 출력 디렉토리 설정
        if output_dir is None:
            base_name = os.path.splitext(os.path.basename(self.xlsx_file))[0]
            output_dir = f"output_{base_name}"
        
        os.makedirs(output_dir, exist_ok=True)
        
        # 4. 전략 실행
        if strategy == 'sequential':
            self._execute_sequential(output_dir, file_info)
        
        elif strategy == 'sheet_parallel':
            self._execute_sheet_parallel(output_dir, config)
        
        elif strategy == 'chunk_parallel':
            self._execute_chunk_parallel(output_dir, config)
        
        elif strategy == 'hybrid':
            self._execute_hybrid(output_dir, config)
        
        total_time = time.time() - start_time
        
        print(f"\n{'#'*70}")
        print(f"✅ 변환 완료!")
        print(f"전략: {strategy}")
        print(f"총 시간: {total_time:.3f}초")
        print(f"출력: {output_dir}/")
        print(f"{'#'*70}\n")
        
        return total_time
    
    def _execute_sequential(self, output_dir, file_info):
        """순차 처리 실행"""
        print("🔄 순차 처리 실행 중...\n")
        
        converter = Xlsx2csv(self.xlsx_file, **self.options)
        
        for sheet_info in file_info['sheets']:
            output_file = os.path.join(output_dir, f"{sheet_info['name']}.csv")
            print(f"  처리: {sheet_info['name']} → {output_file}")
            converter.convert(output_file, sheetid=sheet_info['index'])
    
    def _execute_sheet_parallel(self, output_dir, config):
        """시트 병렬 처리 실행"""
        print(f"⚡ 시트 병렬 처리 실행 중 ({config['num_workers']} 워커)...\n")
        
        converter = Xlsx2csvParallel(self.xlsx_file, **self.options)
        # xlsx2csv_parallel은 자동으로 CPU 수만큼 워커 사용
        converter.convert_parallel(output_dir, verbose=True)
    
    def _execute_chunk_parallel(self, output_dir, config):
        """청크 병렬 처리 실행"""
        print(f"⚡ 청크 병렬 처리 실행 중...\n")
        
        converter = Xlsx2csvChunked(self.xlsx_file, **self.options)
        output_file = os.path.join(output_dir, "output.csv")
        
        converter.convert_chunked(
            output_file,
            chunk_size=config['chunk_size'],
            num_workers=config['num_workers']
        )
    
    def _execute_hybrid(self, output_dir, config):
        """하이브리드 처리 실행"""
        print(f"🚀 하이브리드 처리 실행 중...\n")
        print(f"  - 대용량 시트 {len(config['large_sheets'])}개: 청크 병렬")
        print(f"  - 중소형 시트 {len(config['small_sheets'])}개: 시트 병렬\n")
        
        # 대용량 시트: 청크 병렬
        for sheet_info in config['large_sheets']:
            print(f"📦 대용량 시트 처리: {sheet_info['name']} ({sheet_info['rows']:,}행)")
            
            converter = Xlsx2csvChunked(self.xlsx_file, **self.options)
            output_file = os.path.join(output_dir, f"{sheet_info['name']}.csv")
            
            converter.convert_chunked(
                output_file,
                sheet_index=sheet_info['index'],
                chunk_size=config['chunk_size'],
                num_workers=min(4, cpu_count())
            )
        
        # 중소형 시트: 시트 병렬
        if len(config['small_sheets']) > 0:
            print(f"\n📋 중소형 시트 처리: {len(config['small_sheets'])}개")
            
            # 임시로 순차 처리 (간단한 구현)
            converter = Xlsx2csv(self.xlsx_file, **self.options)
            for sheet_info in config['small_sheets']:
                output_file = os.path.join(output_dir, f"{sheet_info['name']}.csv")
                print(f"  - {sheet_info['name']}")
                converter.convert(output_file, sheetid=sheet_info['index'])


def main():
    """CLI 인터페이스"""
    import argparse
    
    parser = argparse.ArgumentParser(description='xlsx2csv 하이브리드 적응형 변환')
    parser.add_argument('input_file', help='입력 xlsx 파일')
    parser.add_argument('--output-dir', help='출력 디렉토리 (기본: auto)')
    parser.add_argument('--delimiter', default=',', help='CSV 구분자')
    
    args = parser.parse_args()
    
    converter = Xlsx2csvHybrid(args.input_file, delimiter=args.delimiter)
    converter.convert_auto(args.output_dir)


if __name__ == '__main__':
    main()
