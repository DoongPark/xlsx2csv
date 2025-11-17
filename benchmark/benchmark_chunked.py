#!/usr/bin/env python3
"""
청크 기반 병렬 처리 성능 비교 벤치마크
"""

import os
import sys
import time
import subprocess

# 경로 설정: src 폴더를 path에 추가
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
from xlsx2csv import Xlsx2csv
from xlsx2csv_chunked import Xlsx2csvChunked

def benchmark_sequential(xlsx_file, output_file):
    """순차 처리 벤치마크"""
    print(f"\n{'='*60}")
    print("순차 처리 (기준선)")
    print(f"{'='*60}")
    
    start = time.time()
    
    with Xlsx2csv(xlsx_file) as converter:
        converter.convert(output_file)
    
    elapsed = time.time() - start
    
    print(f"⏱️  처리 시간: {elapsed:.3f}초")
    
    # 파일 크기 확인
    if os.path.exists(output_file):
        size = os.path.getsize(output_file) / (1024 * 1024)
        print(f"📊 출력 크기: {size:.2f} MB")
        os.remove(output_file)
    
    return elapsed

def benchmark_chunked(xlsx_file, output_file, chunk_size, num_workers):
    """청크 기반 병렬 처리 벤치마크"""
    print(f"\n{'='*60}")
    print(f"청크 병렬 처리 (청크: {chunk_size:,}행, 워커: {num_workers})")
    print(f"{'='*60}")
    
    converter = Xlsx2csvChunked(xlsx_file)
    elapsed = converter.convert_chunked(
        output_file,
        chunk_size=chunk_size,
        num_workers=num_workers
    )
    
    # 파일 크기 확인
    if os.path.exists(output_file):
        size = os.path.getsize(output_file) / (1024 * 1024)
        print(f"📊 출력 크기: {size:.2f} MB")
        os.remove(output_file)
    
    return elapsed

def run_benchmark_suite(xlsx_file, file_label):
    """전체 벤치마크 스위트 실행"""
    print(f"\n{'#'*70}")
    print(f"# 벤치마크: {file_label}")
    print(f"# 파일: {xlsx_file}")
    print(f"{'#'*70}")
    
    results = {}
    
    # 1. 순차 처리 (기준선)
    output = f"output_bench_{file_label}_seq.csv"
    results['sequential'] = benchmark_sequential(xlsx_file, output)
    
    # 2. 청크 병렬 처리 (다양한 설정)
    configs = [
        (25000, 2),
        (25000, 4),
        (50000, 2),
        (50000, 4),
        (25000, 8),
    ]
    
    for chunk_size, workers in configs:
        output = f"output_bench_{file_label}_c{chunk_size}_w{workers}.csv"
        key = f"chunked_{chunk_size}_{workers}"
        try:
            results[key] = benchmark_chunked(xlsx_file, output, chunk_size, workers)
        except Exception as e:
            print(f"⚠️  에러: {e}")
            results[key] = None
    
    return results

def print_results_table(all_results):
    """결과를 테이블 형식으로 출력"""
    print(f"\n{'='*70}")
    print("📊 성능 비교 요약")
    print(f"{'='*70}\n")
    
    for file_label, results in all_results.items():
        print(f"\n{file_label}:")
        print("-" * 70)
        
        baseline = results.get('sequential', 0)
        
        print(f"{'방법':<30} {'시간(초)':>12} {'속도향상':>12} {'효율성':>12}")
        print("-" * 70)
        
        # 순차 처리
        print(f"{'순차 처리 (기준선)':<30} {baseline:>12.3f} {'1.00x':>12} {'100%':>12}")
        
        # 청크 병렬 처리
        configs = [
            ('chunked_25000_2', '청크 25K, 2 워커'),
            ('chunked_25000_4', '청크 25K, 4 워커'),
            ('chunked_50000_2', '청크 50K, 2 워커'),
            ('chunked_50000_4', '청크 50K, 4 워커'),
            ('chunked_25000_8', '청크 25K, 8 워커'),
        ]
        
        for key, label in configs:
            if key in results and results[key]:
                elapsed = results[key]
                speedup = baseline / elapsed
                # 워커 수 추출
                workers = int(key.split('_')[-1])
                efficiency = (speedup / workers) * 100
                
                print(f"{label:<30} {elapsed:>12.3f} {speedup:>11.2f}x {efficiency:>11.1f}%")
        
        print("-" * 70)

def main():
    """메인 벤치마크 실행"""
    test_files = [
        ('test_data/medium_single_100k.xlsx', '100K행'),
        ('test_data/large_single_200k.xlsx', '200K행'),
    ]
    
    all_results = {}
    
    for xlsx_file, label in test_files:
        if os.path.exists(xlsx_file):
            all_results[label] = run_benchmark_suite(xlsx_file, label)
        else:
            print(f"⚠️  파일 없음: {xlsx_file}")
    
    # 최종 결과 테이블
    print_results_table(all_results)
    
    print(f"\n{'='*70}")
    print("✅ 모든 벤치마크 완료!")
    print(f"{'='*70}\n")

if __name__ == '__main__':
    main()
