from fastapi import FastAPI
from docxtpl import DocxTemplate, InlineImage
from pydantic import BaseModel
from typing import List, Optional
from wwsolution_models import ExperimentProtocol
from io import BytesIO
from fastapi.responses import StreamingResponse
from docx.shared import Mm
import requests

app = FastAPI()

@app.post("/")
async def create_solution(context: ExperimentProtocol):
    # Load the template
    template = DocxTemplate("wwsolution_tpl.docx")

    # Render the template with the context
    template.render(context)

    # Save the document into a BytesIO buffer
    result = BytesIO()
    template.save(result)

    # Rewind the buffer to the beginning
    result.seek(0)

    # Return the result as a StreamingResponse
    return StreamingResponse(result,
                             media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                             headers={"Content-Disposition": "attachment; filename=wwsolution.docx"})
