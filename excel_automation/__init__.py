"""
Excel Automation Package

Package cung cấp các công cụ để đọc, ghi và xử lý file Excel tự động.
"""

__version__ = "1.0.0"
__author__ = "Your Name"

from excel_automation.reader import ExcelReader
from excel_automation.writer import ExcelWriter
from excel_automation.processor import ExcelProcessor
from excel_automation.formatter import ExcelFormatter
from excel_automation.validator import DataValidator, ValidationResult, ValidationError
from excel_automation.validation_rules import (
    ValidationRule,
    RequiredRule,
    TypeRule,
    RangeRule,
    RegexRule,
    LengthRule,
    DateRule,
    UniqueRule,
    InSetRule,
    CustomRule
)
from excel_automation.size_filter import SizeFilterManager
from excel_automation.size_filter_config import SizeFilterConfig
from excel_automation.excel_com_manager import ExcelCOMManager
from excel_automation.utils import get_size_sort_key
from excel_automation.carton_allocation_calculator import (
    CartonAllocationCalculator,
    SizeAllocation,
    CombinedCarton,
    AllocationResult
)

__all__ = [
    "ExcelReader",
    "ExcelWriter",
    "ExcelProcessor",
    "ExcelFormatter",
    "DataValidator",
    "ValidationResult",
    "ValidationError",
    "ValidationRule",
    "RequiredRule",
    "TypeRule",
    "RangeRule",
    "RegexRule",
    "LengthRule",
    "DateRule",
    "UniqueRule",
    "InSetRule",
    "CustomRule",
    "SizeFilterManager",
    "SizeFilterConfig",
    "ExcelCOMManager",
    "get_size_sort_key",
    "CartonAllocationCalculator",
    "SizeAllocation",
    "CombinedCarton",
    "AllocationResult",
]

