
from pathlib import Path
import json

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from io import BytesIO

from core.processing import (
    opcoes_do_arquivo,
    gerar_relatorio,
    analisar_acompanhamento,
    gerar_acompanhamento_zip,
)

app = FastAPI(title="Sistema de Controle de Demandas", version="1.0.0")

BASE_DIR = Path(__file__).resolve().parent.parent
INDEX_HTML = (BASE_DIR / "templates" / "index.html").read_text(encoding="utf-8")


async def _ler_upload(arquivo: UploadFile) -> bytes:
    if not arquivo.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Envie uma planilha Excel (.xlsx ou .xls).")
    conteudo = await arquivo.read()
    if not conteudo:
        raise HTTPException(status_code=400, detail="O arquivo enviado está vazio.")
    return conteudo


@app.get("/", response_class=HTMLResponse)
async def home():
    return INDEX_HTML


@app.get("/api/health")
async def health():
    return {"ok": True}


@app.post("/api/opcoes")
async def opcoes(arquivo: UploadFile = File(...)):
    try:
        conteudo = await _ler_upload(arquivo)
        return JSONResponse(opcoes_do_arquivo(conteudo))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao ler a planilha: {e}")


@app.post("/api/relatorio")
async def relatorio(
    arquivo: UploadFile = File(...),
    prazos: str = Form("[]"),
    prioridade: str = Form("Todos"),
    status: str = Form("[]"),
    interessados: str = Form("[]"),
):
    try:
        conteudo = await _ler_upload(arquivo)
        prazos_list = json.loads(prazos)
        status_list = json.loads(status)
        interessados_list = json.loads(interessados)

        dados, nome, resumo = gerar_relatorio(
            conteudo,
            prazos_list,
            prioridade,
            status_list,
            interessados_list,
        )

        headers = {
            "Content-Disposition": f'attachment; filename="{nome}"',
            "X-Total-Processos": str(resumo["total"]),
        }
        return StreamingResponse(
            BytesIO(dados),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar relatório: {e}")


@app.post("/api/acompanhamento/analisar")
async def acompanhamento_analisar(arquivo: UploadFile = File(...)):
    try:
        conteudo = await _ler_upload(arquivo)
        return JSONResponse(analisar_acompanhamento(conteudo))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao analisar acompanhamento: {e}")


@app.post("/api/acompanhamento/gerar")
async def acompanhamento_gerar(arquivo: UploadFile = File(...)):
    try:
        conteudo = await _ler_upload(arquivo)
        dados, info = gerar_acompanhamento_zip(conteudo)
        headers = {
            "Content-Disposition": f'attachment; filename="{info["nome"]}"',
        }
        return StreamingResponse(
            BytesIO(dados),
            media_type="application/zip",
            headers=headers,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao gerar acompanhamento: {e}")
