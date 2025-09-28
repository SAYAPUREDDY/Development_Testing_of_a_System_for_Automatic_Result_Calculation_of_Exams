import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from typing import List, Dict, Tuple, Any, Optional
from pathlib import Path

class ChartUtils:
    def generate_grade_distribution_chart(students_rows: List[Dict[str, Any]], out_path: Path) -> Optional[str]:
        """
        Build a bar chart of counts per Final_Mark, sorted ascending by numeric value.
        Example: 4 students got 1.3, 1 student got 1.0, etc.
        Saved as PNG at out_path. Returns str path or None.
        """
        # collect final marks
        marks: List[str] = []
        for row in students_rows:
            m = row.get("Final_Mark")
            if m is None:
                continue
            s = str(m).strip().replace(",", ".")
            if s:
                marks.append(s)

        if not marks:
            return None

        # count
        counts: Dict[str, int] = {}
        for s in marks:
            counts[s] = counts.get(s, 0) + 1

        # sort ascending numerically when possible
        def sort_key(k: str):
            try:
                return float(k)
            except:
                return float("inf")

        items = sorted(counts.items(), key=lambda kv: sort_key(kv[0]))

        labels = [k for k, _ in items]
        values = [v for _, v in items]

        # plot
        plt.figure(figsize=(8, 5))
        plt.bar(labels, values)
        plt.title("Grade Distribution (Final Marks)")
        plt.xlabel("Final Mark")
        plt.ylabel("Number of Students")
        plt.tight_layout()
        plt.savefig(out_path)
        plt.close()

        return str(out_path)

