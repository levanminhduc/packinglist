import pandas as pd
from pathlib import Path
import sys
from datetime import datetime, timedelta
import random

sys.path.insert(0, str(Path(__file__).parent.parent))

from config.settings import Settings
from excel_automation.writer import ExcelWriter
from excel_automation.formatter import ExcelFormatter


def create_sample_orders():
    settings = Settings()
    output_path = settings.get_input_path("sample_orders.xlsx")

    valid_sizes = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL']
    colors = ['Red', 'Blue', 'Green', 'Black', 'White', 'Yellow', 'Pink', 'Gray']
    styles = ['T-SHIRT-001', 'POLO-002', 'HOODIE-003', 'JACKET-004', 'PANTS-005']
    buyers = ['Nike', 'Adidas', 'Puma', 'Reebok', 'Under Armour']
    
    base_date = datetime(2025, 12, 1)
    
    data = []
    
    for i in range(1, 11):
        data.append({
            'PO': f'PO{1000000 + i}',
            'Style': random.choice(styles),
            'Color': random.choice(colors),
            'Size': random.choice(valid_sizes),
            'Quantity': random.randint(100, 5000),
            'ShipDate': (base_date + timedelta(days=random.randint(0, 60))).strftime('%Y-%m-%d'),
            'Buyer': random.choice(buyers),
            'Carton': random.randint(5, 50)
        })
    
    data.append({
        'PO': 'INVALID',
        'Style': 'T-SHIRT-001',
        'Color': 'Red',
        'Size': 'M',
        'Quantity': 1000,
        'ShipDate': '2025-12-15',
        'Buyer': 'Nike',
        'Carton': 10
    })
    
    data.append({
        'PO': 'PO1000012',
        'Style': 'AB',
        'Color': 'Blue',
        'Size': 'L',
        'Quantity': 2000,
        'ShipDate': '2025-12-20',
        'Buyer': 'Adidas',
        'Carton': 15
    })
    
    data.append({
        'PO': 'PO1000013',
        'Style': 'POLO-002',
        'Color': '',
        'Size': 'M',
        'Quantity': 1500,
        'ShipDate': '2025-12-25',
        'Buyer': 'Puma',
        'Carton': 12
    })
    
    data.append({
        'PO': 'PO1000014',
        'Style': 'HOODIE-003',
        'Color': 'Black',
        'Size': 'XXXL',
        'Quantity': 3000,
        'ShipDate': '2025-12-30',
        'Buyer': 'Reebok',
        'Carton': 20
    })
    
    data.append({
        'PO': 'PO1000015',
        'Style': 'JACKET-004',
        'Color': 'White',
        'Size': 'L',
        'Quantity': -100,
        'ShipDate': '2025-12-31',
        'Buyer': 'Under Armour',
        'Carton': 8
    })
    
    data.append({
        'PO': 'PO1000016',
        'Style': 'PANTS-005',
        'Color': 'Gray',
        'Size': 'M',
        'Quantity': 150000,
        'ShipDate': '2026-01-05',
        'Buyer': 'Nike',
        'Carton': 10
    })
    
    data.append({
        'PO': 'PO1000017',
        'Style': 'T-SHIRT-001',
        'Color': 'Red',
        'Size': 'S',
        'Quantity': 2000,
        'ShipDate': '31/12/2025',
        'Buyer': 'Adidas',
        'Carton': 15
    })
    
    data.append({
        'PO': 'PO1000018',
        'Style': 'POLO-002',
        'Color': 'Blue',
        'Size': 'M',
        'Quantity': 'ABC',
        'ShipDate': '2026-01-10',
        'Buyer': 'Puma',
        'Carton': 12
    })
    
    data.append({
        'PO': 'PO1000019',
        'Style': 'HOODIE-003',
        'Color': 'Green',
        'Size': 'L',
        'Quantity': 1800,
        'ShipDate': '2026-01-15',
        'Buyer': '',
        'Carton': 14
    })
    
    data.append({
        'PO': 'PO1000020',
        'Style': 'JACKET-004',
        'Color': 'Black',
        'Size': 'XL',
        'Quantity': 2500,
        'ShipDate': '2026-01-20',
        'Buyer': 'Reebok',
        'Carton': 0
    })
    
    for i in range(21, 26):
        data.append({
            'PO': f'PO{1000000 + i}',
            'Style': random.choice(styles),
            'Color': random.choice(colors),
            'Size': random.choice(valid_sizes),
            'Quantity': random.randint(100, 5000),
            'ShipDate': (base_date + timedelta(days=random.randint(0, 60))).strftime('%Y-%m-%d'),
            'Buyer': random.choice(buyers),
            'Carton': random.randint(5, 50)
        })
    
    df = pd.DataFrame(data)
    
    writer = ExcelWriter(str(output_path))
    writer.write_dataframe(df, sheet_name='Orders')
    
    formatter = ExcelFormatter(str(output_path))
    formatter.format_header(sheet_name='Orders', bg_color='366092', font_color='FFFFFF')
    formatter.auto_adjust_column_width(sheet_name='Orders')
    formatter.add_borders(sheet_name='Orders')
    formatter.freeze_panes(sheet_name='Orders', row=1)
    
    print(f"✅ Đã tạo sample data tại: {output_path}")
    print(f"📊 Tổng số dòng: {len(df)}")
    print(f"✓ Valid rows: 10")
    print(f"✗ Invalid rows: 10 (để test validation)")


if __name__ == "__main__":
    create_sample_orders()

