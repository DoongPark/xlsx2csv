#!/usr/bin/env python3
"""
대용량 단일 시트 테스트 파일 생성
청크 기반 병렬 처리 테스트용
"""

import openpyxl
from openpyxl import Workbook
import random
import string
from datetime import datetime, timedelta
import os

def random_string(length=10):
    """랜덤 문자열 생성"""
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def random_date():
    """랜덤 날짜 생성"""
    start = datetime(2020, 1, 1)
    days = random.randint(0, 365 * 4)
    return start + timedelta(days=days)

def create_large_single_sheet(filename, rows=100000, cols=15):
    """
    대용량 단일 시트 xlsx 파일 생성
    
    Args:
        filename: 출력 파일명
        rows: 행 수
        cols: 열 수
    """
    print(f"\n{'='*60}")
    print(f"대용량 단일 시트 파일 생성 중...")
    print(f"파일명: {filename}")
    print(f"크기: {rows:,}행 × {cols}열 = {rows*cols:,}개 셀")
    print(f"{'='*60}\n")
    
    wb = Workbook()
    ws = wb.active
    ws.title = "LargeSheet"
    
    # 헤더 생성
    headers = [f"Column{i+1}" for i in range(cols)]
    ws.append(headers)
    
    # 배치 크기 설정 (메모리 효율)
    batch_size = 1000
    total_batches = rows // batch_size
    
    print("데이터 생성 중...")
    for batch_num in range(total_batches):
        batch_data = []
        
        for _ in range(batch_size):
            row_data = []
            for col_idx in range(cols):
                # 다양한 데이터 타입
                if col_idx % 5 == 0:
                    row_data.append(random_string(8))
                elif col_idx % 5 == 1:
                    row_data.append(random.randint(1000, 9999))
                elif col_idx % 5 == 2:
                    row_data.append(round(random.uniform(0, 1000), 2))
                elif col_idx % 5 == 3:
                    row_data.append(random_date())
                else:
                    row_data.append(random.choice(['Active', 'Inactive', 'Pending', 'Completed']))
            
            batch_data.append(row_data)
        
        # 배치 단위로 추가
        for row_data in batch_data:
            ws.append(row_data)
        
        # 진행률 표시
        progress = (batch_num + 1) / total_batches * 100
        if (batch_num + 1) % 10 == 0:
            print(f"진행률: {progress:.1f}% ({(batch_num + 1) * batch_size:,}/{rows:,}행)")
    
    # 저장
    print("\n파일 저장 중...")
    wb.save(filename)
    
    # 파일 정보 출력
    file_size = os.path.getsize(filename)
    file_size_mb = file_size / (1024 * 1024)
    
    print(f"\n{'='*60}")
    print(f"✅ 파일 생성 완료!")
    print(f"파일명: {filename}")
    print(f"파일 크기: {file_size_mb:.2f} MB")
    print(f"데이터: {rows:,}행 × {cols}열")
    print(f"{'='*60}\n")
    
    return filename, file_size_mb

def main():
    """테스트 파일 생성"""
    output_dir = "test_data"
    os.makedirs(output_dir, exist_ok=True)
    
    # 다양한 크기의 단일 시트 파일 생성
    test_files = [
        ("small_single_50k.xlsx", 50000, 15),
        ("medium_single_100k.xlsx", 100000, 15),
        ("large_single_200k.xlsx", 200000, 15),
    ]
    
    results = []
    for filename, rows, cols in test_files:
        filepath = os.path.join(output_dir, filename)
        fname, size = create_large_single_sheet(filepath, rows, cols)
        results.append((fname, rows, cols, size))
    
    # 요약
    print("\n" + "="*60)
    print("생성된 테스트 파일 요약")
    print("="*60)
    for fname, rows, cols, size in results:
        print(f"{os.path.basename(fname):30} {rows:>8,}행 × {cols:>2}열  {size:>8.2f} MB")
    print("="*60)

if __name__ == "__main__":
    main()
