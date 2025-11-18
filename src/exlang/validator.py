# ============================================================
# exlang.validator: minimal schema checks
# ============================================================

from xml.etree import ElementTree as ET

ALLOWED_TYPES = {"number", "string", "date", "bool"}


def validate_xlang_minimal(root: ET.Element) -> list[str]:
    """
    Perform minimal validation of an exlang document.

    Checks:
      - Root tag is xworkbook
      - xsheet name is optional (auto-generated as Sheet1, Sheet2, etc. if omitted)
      - Auto-generated names must not conflict with explicitly named sheets
      - xrow has r
      - xcell has addr and v
      - xrange has from, to, and fill
      - Optional t attributes use only allowed type names
    """
    errors: list[str] = []

    if root.tag != "xworkbook":
        errors.append(f"Root tag must be 'xworkbook' but found '{root.tag}'")
        return errors

    # Check for collisions between auto-generated and explicit sheet names
    explicit_names = set()
    auto_generated_count = 0
    
    for sheet in root.findall("xsheet"):
        name = sheet.attrib.get("name")
        if name:
            explicit_names.add(name)
        else:
            auto_generated_count += 1
    
    # Check if auto-generated names would conflict with explicit names
    for i in range(1, auto_generated_count + 1):
        auto_name = f"Sheet{i}"
        if auto_name in explicit_names:
            errors.append(
                f"Auto-generated sheet name '{auto_name}' conflicts with explicitly named sheet. "
                f"Either name all sheets or ensure explicit names don't use 'Sheet1', 'Sheet2', etc."
            )

    for sheet in root.findall("xsheet"):

        for xrow in sheet.findall("xrow"):
            if "r" not in xrow.attrib:
                errors.append("xrow missing required attribute 'r'")

        for xcell in sheet.findall("xcell"):
            if "addr" not in xcell.attrib:
                errors.append("xcell missing required attribute 'addr'")
            if "v" not in xcell.attrib:
                errors.append("xcell missing required attribute 'v'")
            t = xcell.attrib.get("t")
            if t is not None and t not in ALLOWED_TYPES:
                errors.append(
                    f"xcell at {xcell.attrib.get('addr', '?')} "
                    f"has invalid type hint t='{t}'"
                )

        for xrange in sheet.findall("xrange"):
            if "from" not in xrange.attrib:
                errors.append("xrange missing required attribute 'from'")
            if "to" not in xrange.attrib:
                errors.append("xrange missing required attribute 'to'")
            if "fill" not in xrange.attrib:
                errors.append("xrange missing required attribute 'fill'")
            t = xrange.attrib.get("t")
            if t is not None and t not in ALLOWED_TYPES:
                errors.append(
                    f"xrange from {xrange.attrib.get('from', '?')} to {xrange.attrib.get('to', '?')} "
                    f"has invalid type hint t='{t}'"
                )

    return errors
