import re
from typing import Optional, Tuple, Dict, Any, List
from functions.question_utils import QuestionUtils

class GradeUtils:

    def map_percentage_to_mark(pct: float, table=None) -> str:
        DEFAULT_GRADING_TABLE = [
            (0,  "5,0"),
            (50, "4,0"),
            (55, "3,7"),
            (60, "3,3"),
            (65, "3,0"),
            (70, "2,7"),
            (75, "2,3"),
            (80, "2,0"),
            (85, "1,7"),
            (90, "1,3"),
            (95, "1,0"),
        ]
        if table is None:
            table = DEFAULT_GRADING_TABLE
        for min_pct, grade in table:
            if pct >= min_pct:
                return grade
        return table[-1][1]



    def run_plausibility_checks(qmap: Dict[str, Any], numeric_achieved: Dict[str, Optional[float]], per_q_seen: Dict[str, int]) -> Tuple[List[str], Dict[str, str]]:
        issues: List[str] = []
        per_q_status: Dict[str, str] = {}
        for qnum, entry in qmap.items():
            max_marks = entry.get("max_marks")
            achieved = numeric_achieved.get(qnum)
            seen = per_q_seen.get(qnum, 1)
            status = QuestionUtils.classify_question_status(achieved, max_marks, seen)
            per_q_status[qnum] = status
            if status == "no_match":
                issues.append(f"{qnum}: no grade detected")
            elif status == "double_match":
                issues.append(f"{qnum}: multiple grades detected ({seen})")
            elif status == "invalid":
                issues.append(f"{qnum}: invalid achieved marks ({achieved}) vs max {max_marks}")

        max_total = sum((entry.get("max_marks") or 0) for entry in qmap.values())
        achieved_sum = sum((v or 0) for v in numeric_achieved.values())
        if achieved_sum > max_total + 1e-6:
            issues.append(f"Achieved total ({achieved_sum}) > Max total ({max_total})")
        return issues, per_q_status
        
    def format_grade_string_and_value(s: Optional[str], max_marks: Optional[int] = None) -> Tuple[str, Optional[float]]:
        """
        Format grade string with respect to max_marks:
        - If max_marks > 9 → keep the input as-is (only replace commas with dots).
        - Otherwise apply normalization rules:
            - "3"   -> "3.0"
            - "45"  -> "4.5"
            - "12"  -> "1.2"
            - "3.5" -> "3.5"
        Returns (formatted_string, numeric_value) or ("", None).
        """
        if s is None:
            return "", None

        s = str(s).strip()
        if s == "":
            return "", None

        s = s.replace(",", ".")

        # --- NEW RULE: If max_marks > 9, keep the raw string (no forced normalization) ---
        if max_marks is not None and max_marks > 9:
            try:
                val = float(s)
                return s, val   # keep as-is (e.g., "12" stays "12")
            except ValueError:
                return s, None

        # --- Normal rules if max_marks <= 9 ---
        if "." in s:
            try:
                val = float(s)
                return f"{val:.1f}", val
            except ValueError:
                return s, None

        if s.isdigit():
            if len(s) == 1:  # e.g., "3" -> "3.0"
                val = float(s)
                return f"{s}.0", val
            else:  # e.g., "45" -> "4.5"
                first, rest = s[0], s[1:]
                formatted = first + "." + rest
                try:
                    val = float(formatted)
                    return (f"{val:.1f}", val) if len(rest) == 1 else (f"{val:.2f}", round(val, 2))
                except ValueError:
                    return s, None

        # fallback
        try:
            val = float(s)
            return f"{val:.1f}", val
        except ValueError:
            return s, None

        