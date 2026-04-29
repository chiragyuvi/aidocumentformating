from functools import lru_cache

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    azure_openai_api_key: str
    azure_openai_endpoint: str
    azure_openai_api_version: str = "2024-10-21"
    model_deployment_name: str

    project_name: str = "Legal Formatter"
    max_upload_size_mb: int = 25
    max_blocks_per_chunk: int = 18
    max_guideline_characters: int = 16000
    heading_color_rgb: str | None = "1F4E79"
    table_header_fill_rgb: str | None = "1F4E79"
    table_header_font_rgb: str | None = "FFFFFF"
    cors_origins: list[str] = Field(default_factory=lambda: ["*"])

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )


@lru_cache
def get_settings() -> Settings:
    return Settings()


settings = get_settings()
