from typing import Optional, Tuple
from pathlib import Path
import re

class OCRPostUtils:

    def normalize_digits(text: Optional[str]) -> Tuple[Optional[int | str], bool]:
        """
        Normalize OCR text characters to digits. 
        Returns (normalized_value, changed_flag). 
        If the result is fully numeric, returns it as an int; otherwise as a string.
        """
        if text is None:
            return None, False

        mapping = {
            "H": "4", "h": "4",
            "O": "0", "o": "0", "Q": "0",
            "l": "1", "I": "1", "|": "1", "ß": "1",
            "Z": "2",
            "E": "3",
            "A": "4",
            "S": "5", "$": "5",
            "G": "6", "b": "6",
            "T": "7",
            "B": "8",
            "g": "9", "q": "9"
        }

        s = str(text)
        normalized = "".join(mapping.get(ch, ch) for ch in s)
        changed = normalized != s

        # Convert to integer if fully numeric
        if normalized.isdigit():
            return int(normalized), changed
        else:
            return normalized, changed

    def make_clickable_link(path: Optional[str], text: str) -> str:
        """
        Make Excel HYPERLINK formula pointing to absolute file:// URI for portability.
        """
        if not path:
            return text or ""
        p = Path(path).resolve()
        uri = p.as_uri()
        # return Excel formula
        return f'=HYPERLINK("{uri}", "{text}")'
    
    def classify_id_field(value: Optional[str]) -> str:
        """Classify plausibility of Matriculation Number / Seat Number."""
        if not value:
            return "invalid"
        s = str(value).strip()
        if re.search(r"[A-Za-z]", s):   # contains letters
            return "normalized"
        if re.fullmatch(r"\d+", s):     # all digits
            return "full"
        return "invalid"
