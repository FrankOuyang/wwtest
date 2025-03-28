from fastapi import FastAPI, UploadFile, File
from docxtpl import DocxTemplate, InlineImage
from pydantic import BaseModel
from typing import List, Optional
from wwsolution_models import ExperimentProtocol
from io import BytesIO
from fastapi.responses import StreamingResponse, HTMLResponse, JSONResponse  # 新增JSONResponse
from docx.shared import Mm
import requests
import pandas as pd  # 新增pandas用于处理Excel
from excel_to_json import excel_to_json  # 新增导入
import tempfile
import os
from fastapi import HTTPException

app = FastAPI()

# 添加文件上传页面
@app.get("/", response_class=HTMLResponse)
async def upload_page():
    return """
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input name="file" type="file">
        <input type="submit">
    </form>
    """

# 文件上传处理
@app.post("/upload")
async def upload_and_generate(file: UploadFile = File(...)):
    try:
        # 创建临时文件保存上传的Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        json_data = excel_to_json(tmp_path)
        
        # 删除临时文件
        os.unlink(tmp_path)
        
        # 转换数量字段为字符串
        if "experimental_groups" in json_data and "rows" in json_data["experimental_groups"]:
            for row in json_data["experimental_groups"]["rows"]:
                if len(row) > 3:  # 确保有第4个元素
                    row[3] = str(row[3])  # 将数量转换为字符串
        
        # 2. 将JSON转换为ExperimentProtocol
        try:
            protocol = ExperimentProtocol(**json_data)
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f"JSON数据转换失败: {str(e)}"
            )
        
        # 3. 调用create_solution生成Word文档
        return await create_solution(protocol)
        
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"处理文件时出错: {str(e)}"
        )

# 文件上传处理2
@app.post("/upload_excel/")
async def upload_excel2(file: UploadFile = File(...)):
    try:
        # 读取上传的Excel文件
        df = pd.read_excel(file.file)
        
        # 转换为JSON格式返回
        data = df.to_dict(orient='records')
        return JSONResponse(content={"status": "success", "data": data})
    
    except Exception as e:
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": str(e)}
        )

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

# 添加Vercel需要的ASGI入口
app = app
