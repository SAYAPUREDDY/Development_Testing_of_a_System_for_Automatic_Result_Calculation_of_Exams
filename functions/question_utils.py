import re
from typing import Optional, Tuple
from typing import List, Dict, Any, Optional


class QuestionUtils:
    def parse_question_headline(text: Optional[str]) -> Tuple[Optional[str], Optional[int]]:
        """
        Parse question headline like '1.Frage ... (10 Punkte)' -> ("Q1", 10)
        """
        if not text:
            return None, None
        m = re.search(r"(\d+)\. ?Frage.*\((\d+)\s*Punkte\)", text)
        if m:
            return f"{int(m.group(1))}", int(m.group(2))
        m2 = re.search(r"(\d+)[^\d]*\((\d+)\s*Punkte\)", text)
        if m2:
            return f"{int(m2.group(1))}", int(m2.group(2))
        return None, None
    
    def classify_question_status(num_achieved: Optional[float], max_marks: Optional[float], seen_count: int) -> str:
        if seen_count == 0:
            return "no_match"
        if seen_count > 1:
            return "double_match"
        if num_achieved is None:
            return "invalid"
        try:
            if max_marks is None:
                return "invalid"
            if num_achieved < 0:
                return "invalid"
            if abs(num_achieved - 0.0) < 1e-9:
                return "zero"
            if abs(num_achieved - max_marks) < 1e-9:
                return "full"
            if num_achieved > max_marks:
                return "invalid"
            if num_achieved < max_marks:
                return "ok"
        except Exception:
            return "invalid"
        return "invalid"
    
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
    
