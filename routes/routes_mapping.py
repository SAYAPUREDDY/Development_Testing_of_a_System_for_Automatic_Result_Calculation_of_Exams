from fastapi import FastAPI
from routes.process_files import router as process_router


def include_routes(app: FastAPI):
    app.include_router(process_router, prefix="/processing", tags=["Process"])
   

