"""Read-only FastMCP server for the O365 License Management database."""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from dotenv import load_dotenv

from db import fetch_all, fetch_one


load_dotenv(Path(__file__).resolve().parent / ".env")

mcp = FastMCP(
    "O365 License MCP",
    instructions=(
        "Read-only access to O365 license assignments, departments, license "
        "types, and cost centers. This server cannot create, update, or delete data."
    ),
    host=os.getenv("MCP_HOST", "127.0.0.1"),
    port=int(os.getenv("MCP_PORT", "20000")),
    stateless_http=True,
    json_response=True,
)

READ_ONLY = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)


def _bounded_limit(limit: int, maximum: int = 500) -> int:
    return max(1, min(limit, maximum))


def _bounded_offset(offset: int) -> int:
    return max(0, offset)


@mcp.tool(annotations=READ_ONLY)
def check_database_health() -> Dict[str, Any]:
    """ตรวจสอบว่า MCP เชื่อมต่อฐานข้อมูล License แบบ read-only ได้หรือไม่"""
    row = fetch_one(
        "SELECT current_database() AS database, "
        "current_user AS db_user, "
        "current_setting('transaction_read_only') AS transaction_read_only"
    )
    return {"status": "ok", **row}


@mcp.tool(annotations=READ_ONLY)
def search_licenses(
    query: Optional[str] = None,
    employee_id: Optional[str] = None,
    department: Optional[str] = None,
    license_name: Optional[str] = None,
    cost_center: Optional[str] = None,
    limit: int = 100,
    offset: int = 0,
) -> Dict[str, Any]:
    """ค้นหารายการ License จากรหัส/ชื่อ/อีเมล/แผนก/ชนิด License/Cost Center

    ทุก filter เป็นตัวเลือก employee_id เป็นการค้นหาตรง ส่วน filter อื่นค้นหา
    แบบบางส่วนและไม่แยกตัวพิมพ์ใหญ่เล็ก ผลลัพธ์สูงสุด 500 รายการต่อครั้ง
    """
    where: List[str] = []
    params: List[Any] = []

    if query and query.strip():
        pattern = f"%{query.strip()}%"
        where.append(
            "(l.emp_id::text ILIKE %s OR l.name_email ILIKE %s "
            "OR l.full_name ILIKE %s OR l.disp_name ILIKE %s "
            "OR l.nick_name ILIKE %s)"
        )
        params.extend([pattern] * 5)
    if employee_id and employee_id.strip():
        where.append("l.emp_id::text = %s")
        params.append(employee_id.strip())
    if department and department.strip():
        pattern = f"%{department.strip()}%"
        where.append("(l.depart_id::text ILIKE %s OR d.depart_name ILIKE %s)")
        params.extend([pattern, pattern])
    if license_name and license_name.strip():
        pattern = f"%{license_name.strip()}%"
        where.append("(l.license_id::text ILIKE %s OR t.license_name ILIKE %s)")
        params.extend([pattern, pattern])
    if cost_center and cost_center.strip():
        pattern = f"%{cost_center.strip()}%"
        where.append("(l.cost_center ILIKE %s OR c.cost_name ILIKE %s)")
        params.extend([pattern, pattern])

    where_sql = " WHERE " + " AND ".join(where) if where else ""
    page_limit = _bounded_limit(limit)
    page_offset = _bounded_offset(offset)
    base_sql = (
        " FROM listLicense_O365 l "
        "LEFT JOIN departments d ON l.depart_id = d.depart_id "
        "LEFT JOIN license_type t ON l.license_id = t.license_id "
        "LEFT JOIN costcenter c ON l.cost_center::text = c.cost_id::text"
    )

    items = fetch_all(
        "SELECT l.emp_id, l.name_email, l.full_name, l.disp_name, "
        "l.nick_name, l.depart_id, d.depart_name, l.cost_center, "
        "c.cost_name, l.license_id, t.license_name"
        + base_sql
        + where_sql
        + " ORDER BY l.emp_id, t.license_name NULLS LAST LIMIT %s OFFSET %s",
        [*params, page_limit, page_offset],
    )
    total_row = fetch_one("SELECT COUNT(*) AS total" + base_sql + where_sql, params)
    return {
        "items": items,
        "total": int(total_row.get("total", 0)),
        "limit": page_limit,
        "offset": page_offset,
        "read_only": True,
    }


@mcp.tool(annotations=READ_ONLY)
def get_employee_licenses(employee_id: str) -> Dict[str, Any]:
    """ดู License ทั้งหมดของพนักงานหนึ่งคนจาก employee ID แบบตรงตัว"""
    employee_id = employee_id.strip()
    if not employee_id:
        raise ValueError("employee_id is required")
    return search_licenses(employee_id=employee_id, limit=500)


@mcp.tool(annotations=READ_ONLY)
def get_license_summary(department: Optional[str] = None) -> Dict[str, Any]:
    """สรุปจำนวน License แยกตามชนิด และเลือกกรองตามชื่อหรือรหัสแผนกได้"""
    where_sql = ""
    params: List[Any] = []
    if department and department.strip():
        pattern = f"%{department.strip()}%"
        where_sql = " WHERE l.depart_id::text ILIKE %s OR d.depart_name ILIKE %s"
        params = [pattern, pattern]

    items = fetch_all(
        "SELECT l.license_id, COALESCE(t.license_name, 'Unknown') AS license_name, "
        "COUNT(*) AS assigned_count "
        "FROM listLicense_O365 l "
        "LEFT JOIN departments d ON l.depart_id = d.depart_id "
        "LEFT JOIN license_type t ON l.license_id = t.license_id"
        + where_sql
        + " GROUP BY l.license_id, t.license_name "
        "ORDER BY assigned_count DESC, license_name",
        params,
    )
    total = sum(int(item["assigned_count"]) for item in items)
    return {"items": items, "total_assignments": total, "read_only": True}


@mcp.tool(annotations=READ_ONLY)
def get_department_summary(limit: int = 100) -> Dict[str, Any]:
    """สรุปจำนวนรายการ License แยกตามแผนก เรียงจากมากไปน้อย"""
    page_limit = _bounded_limit(limit)
    items = fetch_all(
        "SELECT l.depart_id, COALESCE(d.depart_name, 'Unknown') AS depart_name, "
        "COUNT(*) AS assigned_count, COUNT(DISTINCT l.emp_id) AS employee_count "
        "FROM listLicense_O365 l "
        "LEFT JOIN departments d ON l.depart_id = d.depart_id "
        "GROUP BY l.depart_id, d.depart_name "
        "ORDER BY assigned_count DESC, depart_name LIMIT %s",
        [page_limit],
    )
    return {"items": items, "count": len(items), "read_only": True}


@mcp.tool(annotations=READ_ONLY)
def list_reference_data(
    data_type: str,
    search: Optional[str] = None,
    limit: int = 100,
) -> Dict[str, Any]:
    """ดูข้อมูลอ้างอิง โดย data_type ต้องเป็น departments, license_types หรือ cost_centers"""
    definitions = {
        "departments": ("departments", "depart_id", "depart_name"),
        "license_types": ("license_type", "license_id", "license_name"),
        "cost_centers": ("costcenter", "cost_id", "cost_name"),
    }
    if data_type not in definitions:
        raise ValueError(
            "data_type must be departments, license_types, or cost_centers"
        )

    table, id_column, name_column = definitions[data_type]
    params: List[Any] = []
    where_sql = ""
    if search and search.strip():
        pattern = f"%{search.strip()}%"
        where_sql = f" WHERE {id_column}::text ILIKE %s OR {name_column} ILIKE %s"
        params = [pattern, pattern]
    page_limit = _bounded_limit(limit)
    rows = fetch_all(
        f"SELECT {id_column} AS id, {name_column} AS name FROM {table}"
        + where_sql
        + f" ORDER BY {name_column} LIMIT %s",
        [*params, page_limit],
    )
    return {"data_type": data_type, "items": rows, "count": len(rows), "read_only": True}


def main() -> None:
    transport = os.getenv("MCP_TRANSPORT", "stdio").strip().lower()
    if transport not in {"stdio", "streamable-http"}:
        raise ValueError("MCP_TRANSPORT must be 'stdio' or 'streamable-http'")
    mcp.run(transport=transport)


if __name__ == "__main__":
    main()
