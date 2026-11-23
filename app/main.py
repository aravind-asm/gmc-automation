import asyncio
from contextlib import asynccontextmanager, suppress

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app import auto_sync, db
from app.api import health, insights
from app.logging_config import setup_logging

# Before app creation, so anything logged during import is captured too.
setup_logging()


@asynccontextmanager
async def lifespan(_app: FastAPI):
    # Runs once when the server starts (and when a test opens `with TestClient(app)`).
    db.init_db()
    # The automatic daily sync (ADR-012): starts nothing unless AUTO_SYNC_ENABLED is set.
    timer = auto_sync.start()
    yield
    if timer is not None:
        timer.cancel()
        with suppress(asyncio.CancelledError):
            await timer


app = FastAPI(title="GMC Monitoring Insights API", lifespan=lifespan)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3003"], # Updated for GMC Monitoring Insights
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


app.include_router(health.router)
app.include_router(insights.router)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8002)
