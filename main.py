 
from fastapi import FastAPI
from pathlib import Path
import os
import uvicorn
from routes.routes_mapping import include_routes

UPLOAD_DIR = Path("uploads")


os.makedirs(UPLOAD_DIR, exist_ok=True)

app = FastAPI(title="Exam Evaluation", description="Ease of evaluation", version="1.0.0")

include_routes(app)

@app.get("/")
async def root():
    return {"message": "Welcome to the ThinkCompanion, Your Intelligent Documemt Ally!"}

# Run the application with Uvicorn
if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)