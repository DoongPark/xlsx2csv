#!/usr/bin/env python3
"""
복합 xlsx 파일 생성 (다중 시트 + 대용량 시트 포함)
하이브리드 전략 테스트용
"""

import openpyxl
from openpyxl import Workbook
import random
import string
from datetime import datetime, timedelta
import os

def random_string(length=10):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def random_date():
    start = datetime(2020, 1, 1)
    days = random.randint(0, 365 * 4)
    return start + timedelta(days=days)

def create_hybrid_test_file(filename):
    """
    복합 테스트 파일 생성:
    - Sheet1: 대용량 (100,000행)
    - Sheet2-4: 중형 (10,000행)
    - Sheet5-7: 소형 (1,000행)
    """
    print(f"\n{'='*60}")
    print(f"복합 테스트 파일 생성 중...")
    print(f"파일명: {filename}")
    print(f"{'='*60}\n")
    
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거
    
    # 시트 구성
    sheets_config = [
        ("LargeSheet", 100000, 15),
        ("MediumSheet1", 10000, 15),
        ("MediumSheet2", 10000, 15),
        ("MediumSheet3", 10000, 15),
        ("SmallSheet1", 1000, 10),
        ("SmallSheet2", 1000, 10),
        ("SmallSheet3", 1000, 10),
    ]
    
    for sheet_name, rows, cols in sheets_config:
        print(f"시트 생성 중: {sheet_name} ({rows:,}행 × {cols}열)")
        
        ws = wb.create_sheet(title=sheet_name)
        
        # 헤더
        headers = [f"Column{i+1}" for i in range(cols)]
        ws.append(headers)
        
        # 데이터 생성 (배치 처리)
        batch_size = 1000
        num_batches = rows // batch_size
        
        for batch_num in range(num_batches):
            batch_data = []
            for _ in range(batch_size):
                row_data = []
                for col_idx in range(cols):
                    if col_idx % 5 == 0:
                        row_data.append(random_string(8))
                    elif col_idx % 5 == 1:
                        row_data.append(random.randint(1000, 9999))
                    elif col_idx % 5 == 2:
                        row_data.append(round(random.uniform(0, 1000), 2))
                    elif col_idx % 5 == 3:
                        row_data.append(random_date())
                    else:
                        row_data.append(random.choice(['Active', 'Inactive', 'Pending']))
                batch_data.append(row_data)
            
            for row_data in batch_data:
                ws.append(row_data)
            
            if (batch_num + 1) % 10 == 0:
                progress = (batch_num + 1) / num_batches * 100
                print(f"  진행률: {progress:.1f}%")
    
    print("\n파일 저장 중...")
    wb.save(filename)
    
    file_size = os.path.getsize(filename)
    file_size_mb = file_size / (1024 * 1024)
    
    print(f"\n{'='*60}")
    print(f"✅ 복합 파일 생성 완료!")
    print(f"파일명: {filename}")
    print(f"파일 크기: {file_size_mb:.2f} MB")
    print(f"시트 구성:")
    for sheet_name, rows, cols in sheets_config:
        print(f"  - {sheet_name}: {rows:,}행 × {cols}열")
    print(f"{'='*60}\n")
    
    return filename, file_size_mb

def main():
    output_dir = "test_data"
    os.makedirs(output_dir, exist_ok=True)
    
    create_hybrid_test_file(os.path.join(output_dir, "hybrid_test.xlsx"))

if __name__ == "__main__":
    main()
