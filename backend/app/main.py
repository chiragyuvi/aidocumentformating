import uuid
import time
import traceback

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from app.services.azure_client import AzureAIClient
from app.services.doc_processor import DocProcessor

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 🔥 allow all (dev mode)
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ai = AzureAIClient()
processor = DocProcessor()


@app.post("/format-final")
async def format_doc(
    contract: UploadFile = File(...),
    template: UploadFile = File(...)
):
    job_id = f"{int(time.time())}_{uuid.uuid4().hex[:4]}"
    output_path = f"final_output_{job_id}.docx"

    try:
        print("🚀 STARTING", flush=True)

        c_bytes = await contract.read()
        t_bytes = await template.read()

        raw_text = processor.extract_text(c_bytes)
        guideline_text = processor.extract_text(t_bytes)
        tables = processor.extract_tables(c_bytes)

        # 🔥 chunking
        chunk_size = 1000
        chunks = [
            raw_text[i:i + chunk_size]
            for i in range(0, len(raw_text), chunk_size)
        ]

        structured = []

        for i, chunk in enumerate(chunks):
            print(f"Processing {i+1}", flush=True)

            try:
                result = ai.map_to_template(chunk, guideline_text)

                if isinstance(result, list):
                    structured.extend(result)

            except Exception as e:
                print("⚠️ fallback used", flush=True)

                for line in chunk.split("\n"):
                    if line.strip():
                        structured.append({
                            "text": line,
                            "type": "paragraph"
                        })

        doc = processor.build_document(structured, tables)
        doc.save(output_path)

        print(f"✅ DONE: {output_path}", flush=True)

        return FileResponse(output_path)

    except Exception as e:
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))