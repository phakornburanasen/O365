# O365 License FastMCP (read-only)

MCP server สำหรับให้ AI เรียกดูข้อมูลระบบจัดการ O365 License จาก PostgreSQL
โดยไม่มี tool สำหรับเพิ่ม แก้ไข หรือลบข้อมูล

## เครื่องมือที่ AI เรียกใช้ได้

- `check_database_health` ตรวจการเชื่อมต่อและยืนยันสถานะ transaction read-only
- `search_licenses` ค้นหาด้วยชื่อ อีเมล รหัสพนักงาน แผนก License หรือ Cost Center
- `get_employee_licenses` ดู License ตามรหัสพนักงาน
- `get_license_summary` สรุปจำนวนแยกตามชนิด License
- `get_department_summary` สรุปจำนวนแยกตามแผนก
- `list_reference_data` ดูรายการแผนก ชนิด License หรือ Cost Center

## ติดตั้ง

ต้องใช้ Python 3.10 ขึ้นไป (แนะนำ 3.11) จาก PowerShell ที่ project root:

```powershell
py -3.11 -m venv FastMCP\.venv
FastMCP\.venv\Scripts\python.exe -m pip install -r FastMCP\requirements.txt
Copy-Item FastMCP\.env.example FastMCP\.env
```

FastMCP ใช้ connection config ชุดเดียวกับ backend โดยเรียก
`backend/database.py` โดยตรง ค่าฐานข้อมูลจึงอ่านจาก `backend/.env`
และมี default เดียวกับ backend เดิม ไม่ต้องกำหนดค่า PostgreSQL ซ้ำใน
`FastMCP/.env`

เพื่อป้องกันการแก้ข้อมูลในระดับสิทธิ์ฐานข้อมูลด้วย แนะนำให้เปลี่ยนบัญชีใน
`backend/.env` เป็นบัญชีที่ระบบอนุญาตตามความเหมาะสม หรือแยกบัญชี read-only
สำหรับการ deploy MCP ในภายหลัง ตัวอย่างสิทธิ์ read-only:

```sql
CREATE USER license_mcp_reader WITH PASSWORD 'replace-with-a-strong-password';
GRANT CONNECT ON DATABASE postgres TO license_mcp_reader;
GRANT USAGE ON SCHEMA public TO license_mcp_reader;
GRANT SELECT ON TABLE
  listLicense_O365, departments, license_type, costcenter
TO license_mcp_reader;
ALTER USER license_mcp_reader SET default_transaction_read_only = on;
```

ตัว server บังคับ session เป็น read-only อีกชั้นหนึ่ง ถึงแม้ตั้งค่าบัญชี DB ผิดพลาด
ก็ไม่เปิดรับ SQL จาก AI และไม่มี write tool ให้เรียก

## เริ่มใช้งาน

ค่าเริ่มต้นเป็น `stdio` ซึ่งเหมาะกับ AI client ที่รัน MCP เป็น subprocess:

```powershell
FastMCP\.venv\Scripts\python.exe FastMCP\server.py
```

ตัวอย่าง config สำหรับ MCP client:

```json
{
  "mcpServers": {
    "o365-license": {
      "command": "C:\\inetpub\\wwwroot\\O365\\FastMCP\\.venv\\Scripts\\python.exe",
      "args": ["C:\\inetpub\\wwwroot\\O365\\FastMCP\\server.py"]
    }
  }
}
```

ถ้าต้องการ Streamable HTTP ให้ตั้ง `MCP_TRANSPORT=streamable-http` ใน `.env`
แล้วเชื่อมต่อที่ `http://127.0.0.1:20000/mcp` โดยค่าเริ่มต้น bind เฉพาะ localhost

## ตรวจสอบด้วย MCP Inspector

```powershell
npx -y @modelcontextprotocol/inspector `
  FastMCP\.venv\Scripts\python.exe FastMCP\server.py
```

> ไม่ควรเปิด Streamable HTTP ออกอินเทอร์เน็ตโดยตรง หากต้องใช้งานข้ามเครื่อง
> ให้เพิ่ม HTTPS และ authentication ที่ reverse proxy ก่อน
