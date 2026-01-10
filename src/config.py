from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass
class AppConfig:
    output_dir: Path
    sap_system: str
    sap_variant: str

    @staticmethod
    def from_env() -> "AppConfig":
        return AppConfig(
            output_dir=Path(os.getenv("APP_OUTPUT_DIR", Path.cwd() / "output")).resolve(),
            sap_system=os.getenv(
                "SAP_SYSTEM",
                "05 - Africa - Angola - FR3 - Unisup Ecc6 Production",
            ),
            sap_variant=os.getenv("SAP_VARIANT", "CLV-PG2025"),
        )


def ensure_dirs(config: AppConfig) -> None:
    config.output_dir.mkdir(parents=True, exist_ok=True)
