#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
xlsx2csv 성능 벤치마크 도구

다양한 xlsx 파일에 대해 변환 성능을 측정하고 비교합니다.
"""

import os
import sys
import time
import json
import platform
import psutil
import argparse
from datetime import datetime
from pathlib import Path

# xlsx2csv 모듈 import (경로 설정: src 폴더)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
from xlsx2csv import Xlsx2csv


class PerformanceBenchmark:
    """성능 벤치마크 클래스"""
    
    def __init__(self, output_dir="benchmark_results"):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        self.results = []
        self.system_info = self.get_system_info()
    
    def get_system_info(self):
        """시스템 정보 수집"""
        return {
            "platform": platform.platform(),
            "processor": platform.processor(),
            "python_version": sys.version,
            "cpu_count": psutil.cpu_count(logical=False),
            "cpu_count_logical": psutil.cpu_count(logical=True),
            "memory_total_gb": round(psutil.virtual_memory().total / (1024**3), 2),
            "timestamp": datetime.now().isoformat()
        }
    
    def measure_memory_usage(self, func, *args, **kwargs):
        """메모리 사용량 측정"""
        process = psutil.Process()
        
        # 시작 메모리
        mem_before = process.memory_info().rss / (1024 * 1024)  # MB
        
        # 함수 실행
        result = func(*args, **kwargs)
        
        # 종료 메모리
        mem_after = process.memory_info().rss / (1024 * 1024)  # MB
        mem_peak = process.memory_info().rss / (1024 * 1024)  # MB
        
        return result, mem_before, mem_after, mem_peak
    
    def benchmark_file(self, xlsx_file, output_csv=None, num_runs=3, 
                      sheetid=0, **xlsx_options):
        """
        단일 파일 벤치마크
        
        Args:
            xlsx_file: 입력 xlsx 파일 경로
            output_csv: 출력 csv 파일 경로 (None이면 임시 파일)
            num_runs: 실행 횟수
            sheetid: 시트 ID (0 = 모든 시트)
            **xlsx_options: Xlsx2csv 옵션
        
        Returns:
            벤치마크 결과 딕셔너리
        """
        if not os.path.exists(xlsx_file):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {xlsx_file}")
        
        file_size_mb = os.path.getsize(xlsx_file) / (1024 * 1024)
        
        print(f"\n벤치마킹: {os.path.basename(xlsx_file)}")
        print(f"  파일 크기: {file_size_mb:.2f} MB")
        print(f"  실행 횟수: {num_runs}")
        
        times = []
        memory_usages = []
        
        for run in range(num_runs):
            print(f"  실행 {run + 1}/{num_runs}...", end='', flush=True)
            
            # 임시 출력 파일
            if output_csv is None:
                temp_output = os.path.join(
                    self.output_dir, 
                    f"temp_{os.path.basename(xlsx_file)}.csv"
                )
            else:
                temp_output = output_csv
            
            # CPU 시간 측정 시작
            cpu_start = time.process_time()
            wall_start = time.time()
            
            try:
                # 메모리 사용량과 함께 변환 실행
                process = psutil.Process()
                mem_before = process.memory_info().rss / (1024 * 1024)
                
                with Xlsx2csv(xlsx_file, **xlsx_options) as xlsx2csv:
                    # 시트 정보 가져오기
                    num_sheets = len(xlsx2csv.workbook.sheets)
                    
                    # 변환 실행
                    if sheetid == 0:
                        # 모든 시트를 디렉토리로
                        output_dir = temp_output.replace('.csv', '_sheets')
                        if not os.path.exists(output_dir):
                            os.makedirs(output_dir)
                        xlsx2csv.convert(output_dir, sheetid=0)
                    else:
                        xlsx2csv.convert(temp_output, sheetid=sheetid)
                
                mem_after = process.memory_info().rss / (1024 * 1024)
                mem_used = mem_after - mem_before
                
            except Exception as e:
                print(f" 실패: {str(e)}")
                return None
            
            # 시간 측정 종료
            wall_time = time.time() - wall_start
            cpu_time = time.process_time() - cpu_start
            
            times.append({
                'wall_time': wall_time,
                'cpu_time': cpu_time
            })
            memory_usages.append(mem_used)
            
            print(f" 완료 ({wall_time:.2f}초)")
            
            # 임시 파일 정리
            if output_csv is None:
                if os.path.exists(temp_output):
                    os.remove(temp_output)
                if sheetid == 0:
                    output_dir = temp_output.replace('.csv', '_sheets')
                    if os.path.exists(output_dir):
                        import shutil
                        shutil.rmtree(output_dir)
        
        # 통계 계산
        wall_times = [t['wall_time'] for t in times]
        cpu_times = [t['cpu_time'] for t in times]
        
        result = {
            'file': os.path.basename(xlsx_file),
            'file_size_mb': round(file_size_mb, 2),
            'num_sheets': num_sheets,
            'num_runs': num_runs,
            'wall_time_avg': round(sum(wall_times) / len(wall_times), 3),
            'wall_time_min': round(min(wall_times), 3),
            'wall_time_max': round(max(wall_times), 3),
            'cpu_time_avg': round(sum(cpu_times) / len(cpu_times), 3),
            'memory_avg_mb': round(sum(memory_usages) / len(memory_usages), 2),
            'memory_max_mb': round(max(memory_usages), 2),
            'throughput_mb_per_sec': round(file_size_mb / (sum(wall_times) / len(wall_times)), 2),
            'sheetid': sheetid,
            'options': xlsx_options
        }
        
        print(f"  평균 시간: {result['wall_time_avg']:.3f}초")
        print(f"  처리량: {result['throughput_mb_per_sec']:.2f} MB/s")
        print(f"  메모리: {result['memory_avg_mb']:.2f} MB")
        
        return result
    
    def benchmark_directory(self, test_data_dir, num_runs=3, **xlsx_options):
        """
        디렉토리 내 모든 xlsx 파일 벤치마크
        
        Args:
            test_data_dir: 테스트 데이터 디렉토리
            num_runs: 각 파일당 실행 횟수
            **xlsx_options: Xlsx2csv 옵션
        """
        xlsx_files = sorted(Path(test_data_dir).glob("*.xlsx"))
        
        if not xlsx_files:
            print(f"경고: {test_data_dir}에서 xlsx 파일을 찾을 수 없습니다.")
            return
        
        print("=" * 70)
        print(f"벤치마크 시작: {len(xlsx_files)}개 파일")
        print("=" * 70)
        
        results = []
        
        for xlsx_file in xlsx_files:
            result = self.benchmark_file(
                str(xlsx_file),
                num_runs=num_runs,
                sheetid=0,  # 모든 시트
                **xlsx_options
            )
            
            if result:
                results.append(result)
                self.results.append(result)
        
        return results
    
    def save_results(self, filename=None):
        """결과를 JSON 파일로 저장"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"benchmark_{timestamp}.json"
        
        filepath = os.path.join(self.output_dir, filename)
        
        data = {
            'system_info': self.system_info,
            'results': self.results
        }
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        print(f"\n결과 저장됨: {filepath}")
        return filepath
    
    def print_summary(self):
        """결과 요약 출력"""
        if not self.results:
            print("결과가 없습니다.")
            return
        
        print("\n" + "=" * 70)
        print("벤치마크 결과 요약")
        print("=" * 70)
        
        # 시스템 정보
        print("\n[시스템 정보]")
        print(f"  플랫폼: {self.system_info['platform']}")
        print(f"  CPU: {self.system_info['processor']}")
        print(f"  CPU 코어: {self.system_info['cpu_count']} "
              f"(논리: {self.system_info['cpu_count_logical']})")
        print(f"  메모리: {self.system_info['memory_total_gb']} GB")
        print(f"  Python: {self.system_info['python_version'].split()[0]}")
        
        # 결과 테이블
        print("\n[성능 결과]")
        print(f"{'파일명':<35s} {'크기(MB)':>10s} {'시트':>5s} "
              f"{'시간(초)':>10s} {'처리량(MB/s)':>12s} {'메모리(MB)':>10s}")
        print("-" * 90)
        
        for result in sorted(self.results, key=lambda x: x['file_size_mb']):
            print(f"{result['file']:<35s} "
                  f"{result['file_size_mb']:>10.2f} "
                  f"{result['num_sheets']:>5d} "
                  f"{result['wall_time_avg']:>10.3f} "
                  f"{result['throughput_mb_per_sec']:>12.2f} "
                  f"{result['memory_avg_mb']:>10.2f}")
        
        # 통계
        total_time = sum(r['wall_time_avg'] for r in self.results)
        avg_throughput = sum(r['throughput_mb_per_sec'] for r in self.results) / len(self.results)
        
        print("-" * 90)
        print(f"{'전체 평균':<35s} {'':>10s} {'':>5s} "
              f"{total_time:>10.3f} "
              f"{avg_throughput:>12.2f} {'':>10s}")


def main():
    parser = argparse.ArgumentParser(
        description="xlsx2csv 성능 벤치마크 도구"
    )
    parser.add_argument(
        'input',
        nargs='?',
        help='입력 파일 또는 디렉토리 경로'
    )
    parser.add_argument(
        '--test-data-dir',
        default='test_data',
        help='테스트 데이터 디렉토리 (기본값: test_data)'
    )
    parser.add_argument(
        '--output-dir',
        default='benchmark_results',
        help='결과 출력 디렉토리 (기본값: benchmark_results)'
    )
    parser.add_argument(
        '--runs',
        type=int,
        default=3,
        help='각 파일당 실행 횟수 (기본값: 3)'
    )
    parser.add_argument(
        '--sheetid',
        type=int,
        default=0,
        help='변환할 시트 ID (0 = 모든 시트, 기본값: 0)'
    )
    parser.add_argument(
        '--output',
        help='결과 JSON 파일명'
    )
    
    args = parser.parse_args()
    
    benchmark = PerformanceBenchmark(output_dir=args.output_dir)
    
    # 입력에 따라 벤치마크 실행
    if args.input:
        if os.path.isfile(args.input):
            # 단일 파일
            result = benchmark.benchmark_file(
                args.input,
                num_runs=args.runs,
                sheetid=args.sheetid
            )
            if result:
                benchmark.results.append(result)
        elif os.path.isdir(args.input):
            # 디렉토리
            benchmark.benchmark_directory(
                args.input,
                num_runs=args.runs
            )
        else:
            print(f"오류: {args.input}를 찾을 수 없습니다.")
            return 1
    else:
        # 기본 테스트 데이터 디렉토리
        if not os.path.exists(args.test_data_dir):
            print(f"오류: 테스트 데이터 디렉토리를 찾을 수 없습니다: {args.test_data_dir}")
            print("먼저 generate_test_data.py를 실행하여 테스트 데이터를 생성하세요.")
            return 1
        
        benchmark.benchmark_directory(
            args.test_data_dir,
            num_runs=args.runs
        )
    
    # 결과 출력 및 저장
    benchmark.print_summary()
    benchmark.save_results(filename=args.output)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
