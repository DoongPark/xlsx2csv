#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
병렬 vs 순차 성능 비교 벤치마크

원본 xlsx2csv와 병렬 버전의 성능을 비교합니다.
"""

import os
import sys
import time
import json
import psutil
from pathlib import Path
from datetime import datetime

# 경로 설정: src 폴더를 path에 추가
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
from xlsx2csv import Xlsx2csv
from xlsx2csv_parallel import Xlsx2csvParallel
import shutil


class ComparativeBenchmark:
    """순차 vs 병렬 비교 벤치마크"""
    
    def __init__(self, output_dir="benchmark_results"):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        self.results = []
    
    def benchmark_sequential(self, xlsx_file, num_runs=3, **options):
        """순차 처리 벤치마크"""
        print(f"\n[순차 처리] {os.path.basename(xlsx_file)}")
        
        times = []
        memory_usages = []
        
        for run in range(num_runs):
            print(f"  실행 {run + 1}/{num_runs}...", end='', flush=True)
            
            temp_output = os.path.join(self.output_dir, "temp_sequential")
            if not os.path.exists(temp_output):
                os.makedirs(temp_output)
            
            process = psutil.Process()
            mem_before = process.memory_info().rss / (1024 * 1024)
            
            start_time = time.time()
            
            try:
                with Xlsx2csv(xlsx_file, **options) as xlsx2csv:
                    xlsx2csv.convert(temp_output, sheetid=0)
                
                elapsed = time.time() - start_time
                mem_after = process.memory_info().rss / (1024 * 1024)
                mem_used = mem_after - mem_before
                
                times.append(elapsed)
                memory_usages.append(mem_used)
                
                print(f" 완료 ({elapsed:.2f}초)")
            
            except Exception as e:
                print(f" 실패: {str(e)}")
                return None
            
            finally:
                # 정리
                if os.path.exists(temp_output):
                    shutil.rmtree(temp_output)
        
        avg_time = sum(times) / len(times)
        avg_memory = sum(memory_usages) / len(memory_usages)
        
        print(f"  평균 시간: {avg_time:.3f}초")
        print(f"  평균 메모리: {avg_memory:.2f} MB")
        
        return {
            'method': 'sequential',
            'avg_time': avg_time,
            'min_time': min(times),
            'max_time': max(times),
            'avg_memory': avg_memory,
            'times': times
        }
    
    def benchmark_parallel(self, xlsx_file, num_runs=3, num_processes=None, **options):
        """병렬 처리 벤치마크"""
        print(f"\n[병렬 처리] {os.path.basename(xlsx_file)} (프로세스: {num_processes or 'auto'})")
        
        times = []
        memory_usages = []
        
        for run in range(num_runs):
            print(f"  실행 {run + 1}/{num_runs}...", end='', flush=True)
            
            temp_output = os.path.join(self.output_dir, "temp_parallel")
            if not os.path.exists(temp_output):
                os.makedirs(temp_output)
            
            process = psutil.Process()
            mem_before = process.memory_info().rss / (1024 * 1024)
            
            start_time = time.time()
            
            try:
                converter = Xlsx2csvParallel(
                    xlsx_file,
                    num_processes=num_processes,
                    **options
                )
                converter.convert_parallel(temp_output, verbose=False)
                
                elapsed = time.time() - start_time
                mem_after = process.memory_info().rss / (1024 * 1024)
                mem_used = mem_after - mem_before
                
                times.append(elapsed)
                memory_usages.append(mem_used)
                
                print(f" 완료 ({elapsed:.2f}초)")
            
            except Exception as e:
                print(f" 실패: {str(e)}")
                return None
            
            finally:
                # 정리
                if os.path.exists(temp_output):
                    shutil.rmtree(temp_output)
        
        avg_time = sum(times) / len(times)
        avg_memory = sum(memory_usages) / len(memory_usages)
        
        print(f"  평균 시간: {avg_time:.3f}초")
        print(f"  평균 메모리: {avg_memory:.2f} MB")
        
        return {
            'method': 'parallel',
            'num_processes': num_processes,
            'avg_time': avg_time,
            'min_time': min(times),
            'max_time': max(times),
            'avg_memory': avg_memory,
            'times': times
        }
    
    def compare_file(self, xlsx_file, num_runs=3, process_counts=None, **options):
        """
        단일 파일에 대해 순차/병렬 비교
        
        Args:
            xlsx_file: 테스트할 xlsx 파일
            num_runs: 각 방법당 실행 횟수
            process_counts: 테스트할 프로세스 수 리스트 (None이면 [2, 4, 8])
            **options: xlsx2csv 옵션
        """
        file_size_mb = os.path.getsize(xlsx_file) / (1024 * 1024)
        
        # 시트 개수 확인
        with Xlsx2csv(xlsx_file, **options) as xlsx2csv:
            num_sheets = len(xlsx2csv.workbook.sheets)
        
        print("\n" + "=" * 70)
        print(f"파일: {os.path.basename(xlsx_file)}")
        print(f"크기: {file_size_mb:.2f} MB, 시트: {num_sheets}개")
        print("=" * 70)
        
        results = {
            'file': os.path.basename(xlsx_file),
            'file_size_mb': file_size_mb,
            'num_sheets': num_sheets,
            'benchmarks': []
        }
        
        # 순차 처리
        seq_result = self.benchmark_sequential(xlsx_file, num_runs, **options)
        if seq_result:
            results['benchmarks'].append(seq_result)
        
        # 병렬 처리 (다양한 프로세스 수)
        if process_counts is None:
            # CPU 코어 수에 따라 자동 설정
            cpu_count = psutil.cpu_count(logical=True)
            process_counts = [2, 4, cpu_count] if cpu_count >= 4 else [2, cpu_count]
        
        for num_processes in process_counts:
            if num_processes > num_sheets:
                print(f"\n[건너뜀] 프로세스 수({num_processes})가 시트 수({num_sheets})보다 많음")
                continue
            
            par_result = self.benchmark_parallel(
                xlsx_file,
                num_runs,
                num_processes=num_processes,
                **options
            )
            if par_result:
                results['benchmarks'].append(par_result)
        
        # 결과 분석
        self._print_comparison(results)
        self.results.append(results)
        
        return results
    
    def _print_comparison(self, results):
        """비교 결과 출력"""
        print("\n" + "-" * 70)
        print("성능 비교 결과")
        print("-" * 70)
        
        benchmarks = results['benchmarks']
        if not benchmarks:
            return
        
        # 순차 처리 결과 찾기
        seq_result = next((b for b in benchmarks if b['method'] == 'sequential'), None)
        if not seq_result:
            return
        
        seq_time = seq_result['avg_time']
        
        print(f"{'방법':<20s} {'평균 시간(초)':>15s} {'속도 향상':>12s} {'메모리(MB)':>12s}")
        print("-" * 70)
        
        # 순차 처리
        print(f"{'순차 (baseline)':<20s} {seq_time:>15.3f} {'1.00x':>12s} "
              f"{seq_result['avg_memory']:>12.2f}")
        
        # 병렬 처리들
        for bench in benchmarks:
            if bench['method'] == 'parallel':
                speedup = seq_time / bench['avg_time']
                efficiency = speedup / bench['num_processes'] * 100
                
                label = f"병렬 ({bench['num_processes']}코어)"
                print(f"{label:<20s} {bench['avg_time']:>15.3f} "
                      f"{speedup:>11.2f}x "
                      f"{bench['avg_memory']:>12.2f}")
                print(f"{'':20s} {'효율성:':>15s} {efficiency:>10.1f}% {'':>12s}")
    
    def compare_directory(self, test_data_dir, num_runs=3, **options):
        """디렉토리 내 모든 파일 비교"""
        xlsx_files = sorted(Path(test_data_dir).glob("*.xlsx"))
        
        if not xlsx_files:
            print(f"경고: {test_data_dir}에서 xlsx 파일을 찾을 수 없습니다.")
            return
        
        print("=" * 70)
        print(f"비교 벤치마크 시작: {len(xlsx_files)}개 파일")
        print("=" * 70)
        
        for xlsx_file in xlsx_files:
            self.compare_file(str(xlsx_file), num_runs, **options)
    
    def save_results(self, filename=None):
        """결과 저장"""
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"comparison_{timestamp}.json"
        
        filepath = os.path.join(self.output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2)
        
        print(f"\n결과 저장됨: {filepath}")
        return filepath
    
    def print_summary(self):
        """전체 요약 출력"""
        if not self.results:
            return
        
        print("\n" + "=" * 70)
        print("전체 요약")
        print("=" * 70)
        
        print(f"\n{'파일':<30s} {'시트':>6s} {'순차(초)':>12s} "
              f"{'병렬(초)':>12s} {'속도 향상':>12s}")
        print("-" * 70)
        
        total_seq_time = 0
        total_par_time = 0
        
        for result in self.results:
            benchmarks = result['benchmarks']
            seq = next((b for b in benchmarks if b['method'] == 'sequential'), None)
            # 가장 좋은 병렬 결과 선택
            par_results = [b for b in benchmarks if b['method'] == 'parallel']
            par = min(par_results, key=lambda x: x['avg_time']) if par_results else None
            
            if seq and par:
                speedup = seq['avg_time'] / par['avg_time']
                print(f"{result['file']:<30s} {result['num_sheets']:>6d} "
                      f"{seq['avg_time']:>12.3f} {par['avg_time']:>12.3f} "
                      f"{speedup:>11.2f}x")
                
                total_seq_time += seq['avg_time']
                total_par_time += par['avg_time']
        
        if total_seq_time > 0 and total_par_time > 0:
            overall_speedup = total_seq_time / total_par_time
            print("-" * 70)
            print(f"{'전체 평균':<30s} {'':>6s} {total_seq_time:>12.3f} "
                  f"{total_par_time:>12.3f} {overall_speedup:>11.2f}x")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="xlsx2csv 순차 vs 병렬 성능 비교"
    )
    parser.add_argument(
        'input',
        nargs='?',
        help='입력 파일 또는 디렉토리'
    )
    parser.add_argument(
        '--test-data-dir',
        default='test_data',
        help='테스트 데이터 디렉토리'
    )
    parser.add_argument(
        '--runs',
        type=int,
        default=3,
        help='실행 횟수 (기본값: 3)'
    )
    parser.add_argument(
        '--output',
        help='결과 파일명'
    )
    
    args = parser.parse_args()
    
    benchmark = ComparativeBenchmark()
    
    options = {
        'delimiter': ',',
        'outputencoding': 'utf-8'
    }
    
    if args.input:
        if os.path.isfile(args.input):
            benchmark.compare_file(args.input, args.runs, **options)
        elif os.path.isdir(args.input):
            benchmark.compare_directory(args.input, args.runs, **options)
        else:
            print(f"오류: {args.input}를 찾을 수 없습니다.")
            return 1
    else:
        if not os.path.exists(args.test_data_dir):
            print(f"오류: {args.test_data_dir}를 찾을 수 없습니다.")
            return 1
        benchmark.compare_directory(args.test_data_dir, args.runs, **options)
    
    benchmark.print_summary()
    benchmark.save_results(args.output)
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
