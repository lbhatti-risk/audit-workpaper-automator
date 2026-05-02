from dataclasses import dataclass, field
from typing import List
from enum import Enum


class ControlType(Enum):
    ACCESS_PROVISIONING = "access_provisioning"
    ACCESS_DEPROVISIONING = "access_deprovisioning"
    PRIVILEGED_ACCESS = "privileged_access"
    AUTHENTICATION = "authentication"
    DATABASE = "database"
    RECERTIFICATION = "recertification"
    CHANGE_MANAGEMENT_CODE = "change_management_code"
    CHANGE_MANAGEMENT_CONFIG = "change_management_config"


CONTROL_NAMES = {
    ControlType.ACCESS_PROVISIONING: "Access Provisioning (New Starter)",
    ControlType.ACCESS_DEPROVISIONING: "Access Deprovisioning (Leaver)",
    ControlType.PRIVILEGED_ACCESS: "Privileged Access Management",
    ControlType.AUTHENTICATION: "Authentication Controls",
    ControlType.DATABASE: "Database Access Controls",
    ControlType.RECERTIFICATION: "User Access Recertification (UAR)",
    ControlType.CHANGE_MANAGEMENT_CODE: "Change Management - Code Changes",
    ControlType.CHANGE_MANAGEMENT_CONFIG: "Change Management - Configuration Changes",
}

CONTROL_ABBREVS = {
    ControlType.ACCESS_PROVISIONING: "AP",
    ControlType.ACCESS_DEPROVISIONING: "AD",
    ControlType.PRIVILEGED_ACCESS: "PA",
    ControlType.AUTHENTICATION: "AUTH",
    ControlType.DATABASE: "DB",
    ControlType.RECERTIFICATION: "UAR",
    ControlType.CHANGE_MANAGEMENT_CODE: "CM-Code",
    ControlType.CHANGE_MANAGEMENT_CONFIG: "CM-Config",
}


@dataclass
class EvidenceItem:
    filename: str
    file_path: str
    file_type: str
    order: int = 0
    hint: str = ""
    description: str = ""
    key_findings: List[str] = field(default_factory=list)
    aligns_with_process: str = "N/A"
    alignment_explanation: str = ""
    concerns: str = "None noted"


@dataclass
class ProcedureResult:
    number: int
    procedure_text: str
    response: str = ""
    evidence_references: str = ""
    has_gap: bool = False


@dataclass
class ControlDeficiency:
    cd_number: int
    procedure_number: int
    description: str
    severity: str = "Control Deficiency"
    management_response: str = "Pending management response"


@dataclass
class ITGCControl:
    control_type: ControlType
    application: str
    client_name: str
    audit_period: str
    process_description: str
    evidence_items: List[EvidenceItem] = field(default_factory=list)
    procedure_results: List[ProcedureResult] = field(default_factory=list)
    deficiencies: List[ControlDeficiency] = field(default_factory=list)
    conclusion: str = ""

    @property
    def name(self) -> str:
        return CONTROL_NAMES.get(self.control_type, "Unknown Control")

    @property
    def abbrev(self) -> str:
        return CONTROL_ABBREVS.get(self.control_type, "CTRL")
