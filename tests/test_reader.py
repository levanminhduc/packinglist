"""
Unit tests cho ExcelReader.
"""

import pytest
import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_automation import ExcelReader, ExcelWriter


@pytest.fixture
def sample_excel_file(tmp_path):
    """Tạo file Excel mẫu cho testing."""
    file_path = tmp_path / "test_sample.xlsx"
    
    df = pd.DataFrame({
        'Name': ['Alice', 'Bob', 'Charlie'],
        'Age': [25, 30, 35],
        'City': ['Hanoi', 'HCMC', 'Danang']
    })
    
    writer = ExcelWriter(str(file_path))
    writer.write_dataframe(df, sheet_name='TestSheet')
    
    return str(file_path)


class TestExcelReader:
    """Test cases cho ExcelReader class."""
    
    def test_reader_initialization(self, sample_excel_file):
        """Test khởi tạo ExcelReader."""
        reader = ExcelReader(sample_excel_file)
        assert reader.file_path.exists()
    
    def test_reader_file_not_found(self):
        """Test khi file không tồn tại."""
        with pytest.raises(FileNotFoundError):
            ExcelReader("nonexistent_file.xlsx")
    
    def test_read_with_pandas(self, sample_excel_file):
        """Test đọc file với pandas."""
        reader = ExcelReader(sample_excel_file)
        df = reader.read_with_pandas(sheet_name='TestSheet')
        
        assert isinstance(df, pd.DataFrame)
        assert len(df) == 3
        assert 'Name' in df.columns
        assert 'Age' in df.columns
    
    def test_get_sheet_names(self, sample_excel_file):
        """Test lấy danh sách sheet names."""
        reader = ExcelReader(sample_excel_file)
        sheet_names = reader.get_sheet_names()
        
        assert isinstance(sheet_names, list)
        assert 'TestSheet' in sheet_names
    
    def test_read_all_sheets(self, sample_excel_file):
        """Test đọc tất cả sheets."""
        reader = ExcelReader(sample_excel_file)
        all_sheets = reader.read_all_sheets()
        
        assert isinstance(all_sheets, dict)
        assert 'TestSheet' in all_sheets
        assert isinstance(all_sheets['TestSheet'], pd.DataFrame)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])

