import re, difflib
from typing import Dict, Any, List, Tuple, Optional

class PageUtils:
    def fuzzy_contains(text: str, candidates: list[str], cutoff: float = 0.3) -> bool:
        """
        Return True if text is fuzzy-close to any candidate.
        """
        if not text:
            return False
        text = text.lower()
        for cand in candidates:
            score = difflib.SequenceMatcher(None, text, cand.lower()).ratio()
            if score >= cutoff:
                return True
        return False

    def parse_page_marker(text: str) -> Optional[Tuple[int, int]]:
        """
        Parse 'Seite X von Y' patterns with fuzzy closeness.
        """
        if not text:
            return None
        s = text.lower().replace(" ", "")

        # fuzzy check for "seite"
        if not PageUtils.fuzzy_contains(s[:8], ["seite"]):
            return None

        # standard patterns (seiteXvonY / seiteX/Y)
        m = re.search(r"(\d+)\s*von\s*(\d+)", s)
        if m:
            return int(m.group(1)), int(m.group(2))

        m2 = re.search(r"(\d+)\s*/\s*(\d+)", s)
        if m2:
            return int(m2.group(1)), int(m2.group(2))

        # fallback: first 2 numbers in string
        nums = re.findall(r"\d+", s)
        if len(nums) >= 2:
            return int(nums[0]), int(nums[1])

        return None
    
    def page_plausibility_check(page_markers: Dict[int, int], pages_in_group: List[Dict[str, Any]]) -> Tuple[str, bool]:
        """
        page_markers: dict mapping printed_page_number -> total_pages (from OCR text)
        pages_in_group: list of page dicts belonging to this student (each has 'page', the PDF page index)
        Return (check_message, is_ok)
        - If we have at least one printed total_pages (Y), we use the most common total_pages as expected.
        - Build set of printed page numbers seen (keys of page_markers) and compare to expected set.
        - If no markers found -> return ("No page markers found", False) or maybe True? we treat as OK=False to inform user.
        """
        if not page_markers:
            return "No page markers found", False

        # choose expected total as the most common 'total' value
        totals = list(page_markers.values())
        expected_total = max(set(totals), key=totals.count)

        seen_printed = set(page_markers.keys())
        expected_set = set(range(1, expected_total + 1))

        missing = sorted(list(expected_set - seen_printed))
        extra = sorted(list(seen_printed - expected_set))



        # Also check number of pages in the group vs expected_total
        # pages_in_group count might be different from len(seen_printed) if markers missing on some pages
        actual_pages_count = len(pages_in_group)

        messages = []
        ok = True
        if missing:
            if actual_pages_count == expected_total:
                messages.append(f"OCR failed to extract page markers for: {missing}")
            else:
                messages.append(f"YOLO failed to detect pages: {missing}")
            ok = False

        if extra:
            messages.append(f"Unexpected printed pages: {extra}")
            ok = False
        if actual_pages_count < expected_total:
            messages.append(f"Group pages ({actual_pages_count}) < expected total ({expected_total})")
            ok = False
        if actual_pages_count > expected_total:
            messages.append(f"Group pages ({actual_pages_count}) > expected total ({expected_total})")
            ok = False

        if ok:
            return "OK", True
        else:
            return "; ".join(messages), False
