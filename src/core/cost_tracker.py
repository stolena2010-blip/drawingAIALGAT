"""
Cost tracking for API usage
============================
Extracted from customer_extractor_v3_dual.py
"""
from typing import Dict, Tuple
from src.utils.logger import get_logger

logger = get_logger(__name__)


class CostTracker:
    """Track API usage and costs.

    Supports per-call pricing: pass *cost* to :meth:`add_usage` for
    accurate per-stage / per-model costing.  When *cost* is omitted the
    default base-model prices are used (backward compatible).
    """
    def __init__(self, input_price_per_1m: float = 2.50, output_price_per_1m: float = 10.00) -> None:
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.total_files = 0
        self.successful_files = 0
        self._input_price = input_price_per_1m
        self._output_price = output_price_per_1m
        self._accumulated_cost = 0.0  # precise per-call dollar accumulator

    def add_usage(self, input_tokens: int, output_tokens: int,
                  cost: float | None = None) -> float:
        """Accumulate tokens and cost.

        Parameters
        ----------
        cost : float, optional
            Pre-computed dollar cost for this call.  When supplied, the
            tracker uses it directly instead of the base-model prices.

        Returns the dollar cost that was added.
        """
        self.total_input_tokens += input_tokens
        self.total_output_tokens += output_tokens
        if cost is not None:
            self._accumulated_cost += cost
            return cost
        c = (input_tokens / 1_000_000) * self._input_price + \
            (output_tokens / 1_000_000) * self._output_price
        self._accumulated_cost += c
        return c

    def calculate_cost(self) -> Tuple[float, float, float]:
        input_cost = (self.total_input_tokens / 1_000_000) * self._input_price
        output_cost = (self.total_output_tokens / 1_000_000) * self._output_price
        return input_cost, output_cost, self._accumulated_cost
    
    def get_summary(self) -> Dict:
        """Get summary data as dictionary"""
        input_cost, output_cost, total_cost = self.calculate_cost()
        avg_cost = total_cost / self.successful_files if self.successful_files > 0 else 0
        
        return {
            'total_files': self.total_files,
            'successful_files': self.successful_files,
            'input_tokens': self.total_input_tokens,
            'output_tokens': self.total_output_tokens,
            'input_cost': input_cost,
            'output_cost': output_cost,
            'total_cost': total_cost,
            'avg_cost': avg_cost,
            'input_cost_ils': input_cost * 3.7,
            'output_cost_ils': output_cost * 3.7,
            'total_cost_ils': total_cost * 3.7,
            'avg_cost_ils': avg_cost * 3.7
        }
    
    def print_summary(self) -> None:
        input_cost, output_cost, total_cost = self.calculate_cost()
        
        lines = [
            "=" * 70,
            "Cost Summary",
            "=" * 70,
            f"Total files processed: {self.total_files}",
            f"Successful: {self.successful_files}",
            f"Tokens:",
            f"Input:  {self.total_input_tokens:,} tokens",
            f"Output: {self.total_output_tokens:,} tokens",
            f"Costs (Current Model):",
            f"Input:  ${input_cost:.4f} USD ({input_cost * 3.7:.2f})",
            f"Output: ${output_cost:.4f} USD ({output_cost * 3.7:.2f})",
            "",
            f"Total:  ${total_cost:.4f} USD ({total_cost * 3.7:.2f})",
        ]
        
        if self.successful_files > 0:
            avg_cost = total_cost / self.successful_files
            lines.append(f"Average per drawing: ${avg_cost:.6f} USD ({avg_cost * 3.7:.4f})")
        
        lines.append("=" * 70)
        
        for line in lines:
            print(line)          # GUI display
            logger.info(line)    # log file
