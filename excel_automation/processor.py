"""
Module xử lý và transform dữ liệu Excel.
"""

import pandas as pd
from typing import List, Dict, Any, Optional, Callable
import logging

logger = logging.getLogger(__name__)


class ExcelProcessor:
    """Class xử lý và transform dữ liệu Excel."""
    
    def __init__(self):
        logger.info("Khởi tạo ExcelProcessor")
    
    def clean_data(
        self,
        df: pd.DataFrame,
        drop_duplicates: bool = True,
        drop_na: bool = False,
        fill_na: Optional[Any] = None
    ) -> pd.DataFrame:
        """
        Làm sạch dữ liệu.
        
        Args:
            df: DataFrame cần làm sạch
            drop_duplicates: Xóa dòng trùng lặp
            drop_na: Xóa dòng có giá trị null
            fill_na: Giá trị để fill vào chỗ null
            
        Returns:
            DataFrame đã được làm sạch
        """
        try:
            df_clean = df.copy()
            
            if drop_duplicates:
                before = len(df_clean)
                df_clean = df_clean.drop_duplicates()
                logger.info(f"Xóa {before - len(df_clean)} dòng trùng lặp")
            
            if drop_na:
                before = len(df_clean)
                df_clean = df_clean.dropna()
                logger.info(f"Xóa {before - len(df_clean)} dòng có null")
            
            if fill_na is not None:
                df_clean = df_clean.fillna(fill_na)
                logger.info(f"Fill null với giá trị: {fill_na}")
            
            return df_clean
        except Exception as e:
            logger.error(f"Lỗi khi làm sạch dữ liệu: {e}")
            raise
    
    def filter_data(
        self,
        df: pd.DataFrame,
        conditions: Dict[str, Any]
    ) -> pd.DataFrame:
        """
        Lọc dữ liệu theo điều kiện.
        
        Args:
            df: DataFrame cần lọc
            conditions: Dictionary điều kiện {column: value}
            
        Returns:
            DataFrame đã lọc
        """
        try:
            df_filtered = df.copy()
            
            for column, value in conditions.items():
                if column in df_filtered.columns:
                    df_filtered = df_filtered[df_filtered[column] == value]
                    logger.info(f"Lọc {column} = {value}: còn {len(df_filtered)} dòng")
            
            return df_filtered
        except Exception as e:
            logger.error(f"Lỗi khi lọc dữ liệu: {e}")
            raise
    
    def aggregate_data(
        self,
        df: pd.DataFrame,
        group_by: List[str],
        agg_dict: Dict[str, str]
    ) -> pd.DataFrame:
        """
        Tổng hợp dữ liệu theo nhóm.
        
        Args:
            df: DataFrame cần tổng hợp
            group_by: Danh sách cột để group
            agg_dict: Dictionary {column: function} (vd: {'amount': 'sum'})
            
        Returns:
            DataFrame đã tổng hợp
        """
        try:
            df_agg = df.groupby(group_by).agg(agg_dict).reset_index()
            logger.info(f"Tổng hợp theo {group_by}: {len(df_agg)} nhóm")
            return df_agg
        except Exception as e:
            logger.error(f"Lỗi khi tổng hợp dữ liệu: {e}")
            raise
    
    def merge_data(
        self,
        df1: pd.DataFrame,
        df2: pd.DataFrame,
        on: str,
        how: str = 'inner'
    ) -> pd.DataFrame:
        """
        Merge hai DataFrame.
        
        Args:
            df1: DataFrame thứ nhất
            df2: DataFrame thứ hai
            on: Cột để merge
            how: Kiểu merge ('inner', 'left', 'right', 'outer')
            
        Returns:
            DataFrame đã merge
        """
        try:
            df_merged = pd.merge(df1, df2, on=on, how=how)
            logger.info(f"Merge thành công: {len(df_merged)} dòng")
            return df_merged
        except Exception as e:
            logger.error(f"Lỗi khi merge dữ liệu: {e}")
            raise
    
    def pivot_data(
        self,
        df: pd.DataFrame,
        index: str,
        columns: str,
        values: str,
        aggfunc: str = 'sum'
    ) -> pd.DataFrame:
        """
        Tạo pivot table.
        
        Args:
            df: DataFrame nguồn
            index: Cột làm index
            columns: Cột làm columns
            values: Cột chứa giá trị
            aggfunc: Hàm tổng hợp
            
        Returns:
            Pivot table
        """
        try:
            df_pivot = df.pivot_table(
                index=index,
                columns=columns,
                values=values,
                aggfunc=aggfunc
            )
            logger.info(f"Tạo pivot table: {df_pivot.shape}")
            return df_pivot
        except Exception as e:
            logger.error(f"Lỗi khi tạo pivot table: {e}")
            raise
    
    def apply_function(
        self,
        df: pd.DataFrame,
        column: str,
        func: Callable
    ) -> pd.DataFrame:
        """
        Áp dụng function lên một cột.
        
        Args:
            df: DataFrame
            column: Tên cột
            func: Function để áp dụng
            
        Returns:
            DataFrame với cột đã được transform
        """
        try:
            df_result = df.copy()
            df_result[column] = df_result[column].apply(func)
            logger.info(f"Áp dụng function lên cột '{column}'")
            return df_result
        except Exception as e:
            logger.error(f"Lỗi khi áp dụng function: {e}")
            raise
    
    def sort_data(
        self,
        df: pd.DataFrame,
        by: List[str],
        ascending: bool = True
    ) -> pd.DataFrame:
        """
        Sắp xếp dữ liệu.
        
        Args:
            df: DataFrame cần sắp xếp
            by: Danh sách cột để sắp xếp
            ascending: True = tăng dần, False = giảm dần
            
        Returns:
            DataFrame đã sắp xếp
        """
        try:
            df_sorted = df.sort_values(by=by, ascending=ascending)
            logger.info(f"Sắp xếp theo {by}")
            return df_sorted
        except Exception as e:
            logger.error(f"Lỗi khi sắp xếp: {e}")
            raise
    
    def add_calculated_column(
        self,
        df: pd.DataFrame,
        new_column: str,
        formula: str
    ) -> pd.DataFrame:
        """
        Thêm cột tính toán.
        
        Args:
            df: DataFrame
            new_column: Tên cột mới
            formula: Công thức tính (vd: 'col1 + col2')
            
        Returns:
            DataFrame với cột mới
        """
        try:
            df_result = df.copy()
            df_result[new_column] = df_result.eval(formula)
            logger.info(f"Thêm cột '{new_column}' với công thức: {formula}")
            return df_result
        except Exception as e:
            logger.error(f"Lỗi khi thêm cột tính toán: {e}")
            raise

