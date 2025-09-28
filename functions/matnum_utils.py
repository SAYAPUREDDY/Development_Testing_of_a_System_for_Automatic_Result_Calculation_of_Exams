from typing import List, Dict, Any

class MatNumUtils:
    def split_pages_by_matnum(pages: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
        """
        Group pages into exams where each page that contains Mat_num starts a new group.
        If no Mat_num found, return single group with all pages.
        """
        groups: List[List[Dict[str, Any]]] = []
        current: List[Dict[str, Any]] = []
        found_any = False

        for page in pages:
            page_has_mat = False
            for item in page.get("results", []):
                if item.get("label") == "Mat_num" and item.get("text"):
                    page_has_mat = True
                    break

            if page_has_mat:
                found_any = True
                if current:
                    groups.append(current)
                current = [page]
            else:
                current.append(page)

        if current:
            groups.append(current)
        if not found_any:
            return [pages]
        return groups
