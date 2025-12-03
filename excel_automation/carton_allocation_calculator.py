from dataclasses import dataclass, field
from typing import Dict, List, Optional
import logging

from excel_automation.utils import get_size_sort_key

logger = logging.getLogger(__name__)


@dataclass
class SizeAllocation:
    size: str
    total_pcs: int
    full_boxes: int
    full_qty: int
    remainder: int


@dataclass
class CombinedCarton:
    sizes: List[str]
    quantities: Dict[str, int]
    total_pcs: int

    def get_size_label(self, separator: str = "/") -> str:
        return separator.join(self.sizes)

    def is_full(self, items_per_box: int) -> bool:
        return self.total_pcs == items_per_box


@dataclass
class AllocationResult:
    allocations: Dict[str, SizeAllocation]
    combined_cartons: List[CombinedCarton]
    total_full_boxes: int
    total_combined_boxes: int
    items_per_box: int

    @property
    def total_boxes(self) -> int:
        return self.total_full_boxes + self.total_combined_boxes


class CartonAllocationCalculator:

    def __init__(self, items_per_box: int):
        if items_per_box <= 0:
            raise ValueError("items_per_box phải > 0")
        self.items_per_box = items_per_box
        logger.info(f"Khởi tạo CartonAllocationCalculator với items_per_box={items_per_box}")

    def calculate_allocation(self, size: str, total_pcs: int) -> SizeAllocation:
        if total_pcs < 0:
            raise ValueError(f"total_pcs cho size {size} phải >= 0")

        full_boxes = total_pcs // self.items_per_box
        remainder = total_pcs % self.items_per_box
        full_qty = full_boxes * self.items_per_box

        return SizeAllocation(
            size=size,
            total_pcs=total_pcs,
            full_boxes=full_boxes,
            full_qty=full_qty,
            remainder=remainder
        )

    def calculate_all_allocations(
        self,
        size_quantities: Dict[str, int]
    ) -> Dict[str, SizeAllocation]:
        allocations: Dict[str, SizeAllocation] = {}

        for size, total_pcs in size_quantities.items():
            allocations[size] = self.calculate_allocation(size, total_pcs)
            logger.debug(
                f"Size {size}: {total_pcs} pcs -> "
                f"{allocations[size].full_boxes} thùng nguyên, "
                f"{allocations[size].remainder} dư"
            )

        return allocations

    def calculate_combined_cartons(
        self,
        allocations: Dict[str, SizeAllocation]
    ) -> List[CombinedCarton]:
        remainders: List[tuple] = []

        for size, alloc in allocations.items():
            if alloc.remainder > 0:
                remainders.append((size, alloc.remainder))

        remainders.sort(key=lambda x: get_size_sort_key(x[0]))

        combined_cartons: List[CombinedCarton] = []
        current_sizes: List[str] = []
        current_quantities: Dict[str, int] = {}
        current_total = 0

        for size, qty in remainders:
            remaining_qty = qty

            while remaining_qty > 0:
                space_left = self.items_per_box - current_total
                take = min(remaining_qty, space_left)

                if size in current_quantities:
                    current_quantities[size] += take
                else:
                    current_sizes.append(size)
                    current_quantities[size] = take

                current_total += take
                remaining_qty -= take

                if current_total == self.items_per_box:
                    combined_cartons.append(CombinedCarton(
                        sizes=current_sizes.copy(),
                        quantities=current_quantities.copy(),
                        total_pcs=current_total
                    ))
                    current_sizes = []
                    current_quantities = {}
                    current_total = 0

        if current_total > 0:
            combined_cartons.append(CombinedCarton(
                sizes=current_sizes.copy(),
                quantities=current_quantities.copy(),
                total_pcs=current_total
            ))

        logger.info(f"Đã tạo {len(combined_cartons)} thùng ghép")
        return combined_cartons

    def get_full_result(
        self,
        size_quantities: Dict[str, int]
    ) -> AllocationResult:
        logger.info(f"Tính toán phân bổ cho {len(size_quantities)} sizes")

        allocations = self.calculate_all_allocations(size_quantities)
        combined_cartons = self.calculate_combined_cartons(allocations)

        total_full_boxes = sum(a.full_boxes for a in allocations.values())

        return AllocationResult(
            allocations=allocations,
            combined_cartons=combined_cartons,
            total_full_boxes=total_full_boxes,
            total_combined_boxes=len(combined_cartons),
            items_per_box=self.items_per_box
        )

