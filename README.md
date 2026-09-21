# AI Assistant

## MSB AI Hackathon 2026 - GreenNode Team

### Menu Structure
```
I. Doi soat du lieu
   I.1. Doi soat du lieu dien
   I.2. Doi soat du lieu nuoc uong

II. Phan bo chi phi
   II.1. Phan bo chi phi dien
   II.2. Phan bo chi phi nuoc uong

III. Quan ly chi phi (coming soon)
```

### Features
- **I. Doi soat:** Compare 2 files by So PGH (9-digit ticket number)
  - Results: Dung / Sai / Khong co ket qua
  - Output: Excel with result column added
- **II. Phan bo:** Update Excel template with PDF/Excel/Image/API data
  - Match by Ma KH (customer code)
  - Shift dates by month
  - Split quantities for duplicate codes
  - Skip "di di" rows

### Input Formats
- PDF (bank statement, sales detail, delivery ticket)
- Excel (.xlsx, .xls)
- Image (JPG, PNG, GIF) - via OCR
- API (JSON)

### Files
- `app.py` - Desktop GUI app (tkinter)
- `web_app.py` - Web API (FastAPI) for cloud deployment
- `update_excel_auto.py` - CLI script
- `run_app.bat` - Windows launcher
- `Dockerfile` - Docker build for AgentBase
- `rules/` - Rule documentation
- `agent/` - AI agent config

### Deploy
- Runtime: GreenNode AgentBase
- Image: vcr.vngcloud.vn/111480-abp114532/phan-bo-chi-phi-dien:latest
- Console: https://aiplatform.console.vngcloud.vn/agent-runtime?tab=runtime

### API Endpoints
- `POST /phan-bo` - Phan bo chi phi
- `POST /phan-bo/download` - Download result Excel
- `POST /doi-soat` - Doi soat du lieu
- `POST /doi-soat/download` - Download result Excel
- `GET /health` - Health check
- `GET /docs` - API documentation (Swagger)
