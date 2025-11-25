from abc import ABC, abstractmethod
from typing import Any, Optional, List, Callable
import re
from datetime import datetime
import pandas as pd


class ValidationRule(ABC):
    
    def __init__(self, column: str, error_message: Optional[str] = None):
        self.column = column
        self.error_message = error_message
    
    @abstractmethod
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        pass
    
    def get_error_message(self, value: Any, default_message: str) -> str:
        return self.error_message or default_message


class RequiredRule(ValidationRule):
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value) or value is None or (isinstance(value, str) and value.strip() == ""):
            message = self.get_error_message(value, f"Trường '{self.column}' là bắt buộc")
            return False, message
        return True, None


class TypeRule(ValidationRule):
    
    def __init__(self, column: str, expected_type: type, error_message: Optional[str] = None):
        super().__init__(column, error_message)
        self.expected_type = expected_type
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        try:
            if self.expected_type == int:
                int(value)
            elif self.expected_type == float:
                float(value)
            elif self.expected_type == str:
                str(value)
            return True, None
        except (ValueError, TypeError):
            message = self.get_error_message(
                value, 
                f"Giá trị '{value}' không đúng kiểu {self.expected_type.__name__}"
            )
            return False, message


class RangeRule(ValidationRule):
    
    def __init__(
        self, 
        column: str, 
        min_value: Optional[float] = None, 
        max_value: Optional[float] = None,
        error_message: Optional[str] = None
    ):
        super().__init__(column, error_message)
        self.min_value = min_value
        self.max_value = max_value
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        try:
            num_value = float(value)
            
            if self.min_value is not None and num_value < self.min_value:
                message = self.get_error_message(
                    value,
                    f"Giá trị {value} nhỏ hơn giá trị tối thiểu {self.min_value}"
                )
                return False, message
            
            if self.max_value is not None and num_value > self.max_value:
                message = self.get_error_message(
                    value,
                    f"Giá trị {value} lớn hơn giá trị tối đa {self.max_value}"
                )
                return False, message
            
            return True, None
        except (ValueError, TypeError):
            message = self.get_error_message(value, f"Giá trị '{value}' không phải là số")
            return False, message


class RegexRule(ValidationRule):
    
    def __init__(self, column: str, pattern: str, error_message: Optional[str] = None):
        super().__init__(column, error_message)
        self.pattern = pattern
        self.regex = re.compile(pattern)
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        str_value = str(value)
        if not self.regex.match(str_value):
            message = self.get_error_message(
                value,
                f"Giá trị '{value}' không khớp với định dạng yêu cầu"
            )
            return False, message
        
        return True, None


class LengthRule(ValidationRule):
    
    def __init__(
        self, 
        column: str, 
        min_length: Optional[int] = None, 
        max_length: Optional[int] = None,
        error_message: Optional[str] = None
    ):
        super().__init__(column, error_message)
        self.min_length = min_length
        self.max_length = max_length
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        str_value = str(value)
        length = len(str_value)
        
        if self.min_length is not None and length < self.min_length:
            message = self.get_error_message(
                value,
                f"Độ dài '{length}' nhỏ hơn độ dài tối thiểu {self.min_length}"
            )
            return False, message
        
        if self.max_length is not None and length > self.max_length:
            message = self.get_error_message(
                value,
                f"Độ dài '{length}' lớn hơn độ dài tối đa {self.max_length}"
            )
            return False, message
        
        return True, None


class DateRule(ValidationRule):
    
    def __init__(
        self, 
        column: str, 
        date_format: str = "%Y-%m-%d",
        error_message: Optional[str] = None
    ):
        super().__init__(column, error_message)
        self.date_format = date_format
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        if isinstance(value, (datetime, pd.Timestamp)):
            return True, None
        
        try:
            datetime.strptime(str(value), self.date_format)
            return True, None
        except ValueError:
            message = self.get_error_message(
                value,
                f"Giá trị '{value}' không đúng định dạng ngày {self.date_format}"
            )
            return False, message


class UniqueRule(ValidationRule):
    
    def __init__(self, column: str, error_message: Optional[str] = None):
        super().__init__(column, error_message)
        self.seen_values = set()
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        str_value = str(value)
        if str_value in self.seen_values:
            message = self.get_error_message(
                value,
                f"Giá trị '{value}' bị trùng lặp"
            )
            return False, message
        
        self.seen_values.add(str_value)
        return True, None
    
    def reset(self):
        self.seen_values.clear()


class InSetRule(ValidationRule):
    
    def __init__(
        self, 
        column: str, 
        allowed_values: List[Any],
        case_sensitive: bool = True,
        error_message: Optional[str] = None
    ):
        super().__init__(column, error_message)
        self.allowed_values = allowed_values
        self.case_sensitive = case_sensitive
        
        if not case_sensitive:
            self.allowed_values_lower = [str(v).lower() for v in allowed_values]
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        if self.case_sensitive:
            if value not in self.allowed_values:
                message = self.get_error_message(
                    value,
                    f"Giá trị '{value}' không nằm trong danh sách cho phép: {self.allowed_values}"
                )
                return False, message
        else:
            str_value = str(value).lower()
            if str_value not in self.allowed_values_lower:
                message = self.get_error_message(
                    value,
                    f"Giá trị '{value}' không nằm trong danh sách cho phép: {self.allowed_values}"
                )
                return False, message
        
        return True, None


class CustomRule(ValidationRule):
    
    def __init__(
        self, 
        column: str, 
        validation_func: Callable[[Any, Optional[dict]], bool],
        error_message: Optional[str] = None
    ):
        super().__init__(column, error_message)
        self.validation_func = validation_func
    
    def validate(self, value: Any, row_data: Optional[dict] = None) -> tuple[bool, Optional[str]]:
        if pd.isna(value):
            return True, None
        
        try:
            is_valid = self.validation_func(value, row_data)
            if not is_valid:
                message = self.get_error_message(
                    value,
                    f"Giá trị '{value}' không hợp lệ theo quy tắc tùy chỉnh"
                )
                return False, message
            return True, None
        except Exception as e:
            message = self.get_error_message(
                value,
                f"Lỗi khi kiểm tra giá trị '{value}': {str(e)}"
            )
            return False, message

