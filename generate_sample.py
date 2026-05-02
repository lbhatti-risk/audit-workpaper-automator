"""
Generate a sample workpaper using synthetic data — no API key required.
Run:  python generate_sample.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from src.models import (
    ControlType, ITGCControl, EvidenceItem, ProcedureResult, ControlDeficiency
)
from src.controls import TAILORED_PROCEDURES
from src.workpaper import generate_workpaper

OUTPUT = "sample_output/Acme_Corp_Oracle_EBS_ITGC_Workpaper_FY2025_SAMPLE.xlsx"


def _make_access_provisioning() -> ITGCControl:
    ctrl = ITGCControl(
        control_type=ControlType.ACCESS_PROVISIONING,
        application="Oracle EBS",
        client_name="Acme Corporation",
        audit_period="FY2025",
        process_description=(
            "When a new employee joins Acme Corporation, the line manager submits an IT Access Request "
            "Form via the ServiceNow portal specifying the required Oracle EBS modules and roles. "
            "The IT Helpdesk reviews the request and confirms alignment with the employee's job description. "
            "The request is then routed to the IT Security team for approval. Upon approval, the Oracle "
            "Systems Administrator provisions the access within 2 business days. The new user receives an "
            "automated email notification confirming their login credentials and access scope. "
            "A copy of the provisioned access is retained in ServiceNow as the audit trail."
        ),
    )

    procedures = TAILORED_PROCEDURES[ControlType.ACCESS_PROVISIONING]

    ctrl.evidence_items = [
        EvidenceItem(
            filename="SN_Ticket_INC0012345.png",
            file_path="",
            file_type="image",
            order=1,
            hint="ServiceNow access request ticket",
            description=(
                "The screenshot depicts ServiceNow incident ticket INC0012345, dated 14 January 2025, "
                "submitted by J. Smith (line manager) on behalf of new starter E. Johnson. The ticket "
                "records the requested Oracle EBS roles (AR Clerk, GL Viewer) and references the "
                "employee's offer letter as supporting documentation."
            ),
            key_findings=[
                "Formal access request submitted via ServiceNow prior to provisioning.",
                "Requested roles (AR Clerk, GL Viewer) documented within the ticket.",
                "Request date (14 Jan 2025) precedes the employee start date (20 Jan 2025).",
            ],
            aligns_with_process="Yes",
            alignment_explanation="Request was submitted via the defined ServiceNow portal by the line manager.",
            concerns="None noted",
        ),
        EvidenceItem(
            filename="Manager_Approval_Email.png",
            file_path="",
            file_type="image",
            order=2,
            hint="Line manager approval email",
            description=(
                "The screenshot depicts an email dated 15 January 2025 from J. Smith (Line Manager) "
                "to the IT Security team approving the Oracle EBS access request for E. Johnson. "
                "The email references ticket INC0012345 and confirms the business justification for "
                "the requested roles."
            ),
            key_findings=[
                "Written line manager approval obtained before access was provisioned.",
                "Approval references the original ServiceNow ticket number.",
                "Approval provided 5 days before the employee start date.",
            ],
            aligns_with_process="Yes",
            alignment_explanation="Documented line manager approval aligns with the stated process.",
            concerns="None noted",
        ),
        EvidenceItem(
            filename="Oracle_User_Provisioning_Log.xlsx",
            file_path="",
            file_type="excel",
            order=3,
            hint="Oracle user provisioning log extract",
            description=(
                "The spreadsheet extract from the Oracle EBS user provisioning log shows that user "
                "account 'EJOHNSON' was created on 20 January 2025 with the roles AR Clerk and "
                "GL Viewer assigned. The provisioning was completed within the 2 business day SLA "
                "following the approval dated 15 January 2025."
            ),
            key_findings=[
                "Account created on 20 Jan 2025 — within the 2 business day SLA.",
                "Roles provisioned (AR Clerk, GL Viewer) match those requested in INC0012345.",
                "Provisioning log does not contain a field for SoD conflict check.",
            ],
            aligns_with_process="Partial",
            alignment_explanation=(
                "Access was provisioned timely and roles match the request; however, the log does "
                "not evidence a segregation of duties conflict check was performed."
            ),
            concerns=(
                "The provisioning log does not include a field or indicator confirming that a "
                "segregation of duties (SoD) conflict review was completed prior to provisioning."
            ),
        ),
    ]

    ctrl.procedure_results = [
        ProcedureResult(
            number=i + 1,
            procedure_text=proc,
            response=response,
            evidence_references=ev_ref,
            has_gap=gap,
        )
        for i, (proc, response, ev_ref, gap) in enumerate([
            (
                procedures[0],
                "The Access Provisioning Policy (v3.2, approved January 2024) was obtained and inspected. "
                "The policy covers Oracle EBS and specifies the ServiceNow portal as the mandatory request channel. "
                "The policy is signed by the CISO and is within its annual review cycle.",
                "Not evidenced (policy provided separately by management)",
                False,
            ),
            (
                procedures[1],
                "Management provided a complete population of 47 new user access requests provisioned during FY2025. "
                "Completeness was confirmed by agreeing the count to the ServiceNow provisioning report.",
                "Not evidenced (population obtained from management)",
                False,
            ),
            (
                procedures[2],
                "For the sample item tested (INC0012345 — E. Johnson), a formal ServiceNow access request "
                "(INC0012345) was confirmed to have been submitted on 14 January 2025, prior to the employee "
                "start date of 20 January 2025 and prior to access being provisioned.",
                "Evidence 1 (SN_Ticket_INC0012345.png)",
                False,
            ),
            (
                procedures[3],
                "Written approval from J. Smith (line manager) was confirmed via email dated 15 January 2025, "
                "referencing ticket INC0012345. Approval was obtained prior to the access being provisioned on "
                "20 January 2025.",
                "Evidence 2 (Manager_Approval_Email.png)",
                False,
            ),
            (
                procedures[4],
                "The roles provisioned (AR Clerk, GL Viewer) were confirmed to match those requested in "
                "INC0012345 and are consistent with the employee's role as Accounts Receivable Clerk per "
                "their job description on file.",
                "Evidence 1, Evidence 3",
                False,
            ),
            (
                procedures[5],
                "Access was provisioned on 20 January 2025, 3 business days after approval was granted on "
                "15 January 2025. The policy stipulates a 2 business day SLA; however, 20 January 2025 "
                "was the employee's first day and provisioning was completed before the start of business. "
                "No SLA breach was identified.",
                "Evidence 2, Evidence 3",
                False,
            ),
            (
                procedures[6],
                "The Oracle EBS provisioning log does not include a field or indicator confirming that a "
                "segregation of duties (SoD) conflict check was performed prior to provisioning. Management "
                "confirmed verbally that an automated SoD tool is used; however, no documentary evidence of "
                "the SoD review for the sampled request was provided.",
                "Evidence 3 (Oracle_User_Provisioning_Log.xlsx)",
                True,
            ),
        ])
    ]

    ctrl.deficiencies = [
        ControlDeficiency(
            cd_number=1,
            procedure_number=7,
            description=(
                "The Oracle EBS access provisioning process does not retain documented evidence that a "
                "segregation of duties (SoD) conflict check was performed prior to access being granted. "
                "The stated process indicates that SoD conflicts are considered during provisioning; "
                "however, the provisioning log for the sampled request (INC0012345) contains no field "
                "or attachment evidencing this review was completed."
            ),
            severity="Control Deficiency",
            management_response="Pending management response",
        )
    ]

    ctrl.conclusion = (
        "Based on the evidence reviewed, the access provisioning control for Oracle EBS is "
        "substantially designed and implemented in accordance with the stated process. Access requests "
        "are submitted via ServiceNow, manager approval is documented, and access is provisioned "
        "within the defined SLA with roles commensurate with the user's responsibilities. However, "
        "one control deficiency (CD/W-1) has been raised as the process does not retain documentary "
        "evidence of segregation of duties conflict checks performed at the time of provisioning."
    )

    return ctrl


def _make_change_management() -> ITGCControl:
    ctrl = ITGCControl(
        control_type=ControlType.CHANGE_MANAGEMENT_CODE,
        application="Oracle EBS",
        client_name="Acme Corporation",
        audit_period="FY2025",
        process_description=(
            "All code changes to Oracle EBS are initiated via a Change Request (CR) ticket in Jira. "
            "The developer raises the CR, attaches the technical specification, and assigns it to the "
            "Change Advisory Board (CAB) for review. CAB meets weekly and approves or rejects changes. "
            "Approved changes are deployed to the UAT environment by the development team, where the "
            "business owner performs user acceptance testing. Upon sign-off, the change is promoted to "
            "production by the Release Manager — a role held separately from the development team, "
            "ensuring segregation of duties. Emergency changes follow an expedited process with "
            "retrospective CAB approval required within 3 business days."
        ),
    )

    procedures = TAILORED_PROCEDURES[ControlType.CHANGE_MANAGEMENT_CODE]

    ctrl.evidence_items = [
        EvidenceItem(
            filename="Jira_CR_2025-0147.png",
            file_path="",
            file_type="image",
            order=1,
            hint="Jira change request ticket for CR-2025-0147",
            description=(
                "The screenshot depicts Jira Change Request CR-2025-0147, raised on 3 March 2025 by "
                "developer B. Patel. The ticket describes a fix to the Oracle EBS AR module invoice "
                "calculation logic, includes a technical specification attachment, and shows status "
                "'CAB Approved' with approval date 10 March 2025."
            ),
            key_findings=[
                "Formal Jira CR raised prior to development commencing.",
                "Technical specification attached to the ticket.",
                "CAB approval documented within the ticket on 10 March 2025.",
            ],
            aligns_with_process="Yes",
            alignment_explanation="CR follows the defined Jira-based change request process with CAB approval.",
            concerns="None noted",
        ),
        EvidenceItem(
            filename="UAT_Signoff_CR-2025-0147.pdf",
            file_path="",
            file_type="pdf",
            order=2,
            hint="UAT sign-off document",
            description=(
                "The PDF document is the UAT sign-off form for CR-2025-0147, signed by business owner "
                "C. Williams on 17 March 2025. The form confirms that the AR invoice calculation "
                "fix was tested in the UAT environment and the results were satisfactory."
            ),
            key_findings=[
                "UAT testing completed by the business owner (C. Williams) in non-production.",
                "Sign-off dated 17 March 2025, prior to production deployment on 20 March 2025.",
                "UAT form references the Jira CR number for traceability.",
            ],
            aligns_with_process="Yes",
            alignment_explanation="UAT completed by the business owner prior to production deployment as per process.",
            concerns="None noted",
        ),
        EvidenceItem(
            filename="Production_Deployment_Log.xlsx",
            file_path="",
            file_type="excel",
            order=3,
            hint="Production deployment log",
            description=(
                "The spreadsheet extract from the production deployment log shows that CR-2025-0147 "
                "was deployed to production on 20 March 2025 at 22:14 by Release Manager D. Chen. "
                "Developer B. Patel is not recorded as having production deployment access."
            ),
            key_findings=[
                "Production deployment performed by Release Manager (D. Chen), not the developer.",
                "Deployment date (20 Mar 2025) follows UAT sign-off (17 Mar 2025).",
                "Segregation of duties between developer and Release Manager confirmed.",
            ],
            aligns_with_process="Yes",
            alignment_explanation="Deployment by Release Manager with developer excluded confirms SoD.",
            concerns="None noted",
        ),
    ]

    ctrl.procedure_results = [
        ProcedureResult(
            number=i + 1,
            procedure_text=proc,
            response=response,
            evidence_references=ev_ref,
            has_gap=gap,
        )
        for i, (proc, response, ev_ref, gap) in enumerate([
            (
                procedures[0],
                "The Change Management Policy (v2.1, approved April 2024) was obtained and inspected. "
                "The policy covers code changes to Oracle EBS and defines the full change lifecycle "
                "including Jira CR, CAB approval, UAT, and production deployment via the Release Manager.",
                "Not evidenced (policy obtained separately)",
                False,
            ),
            (
                procedures[1],
                "Management provided the complete population of 312 code changes deployed to production "
                "during FY2025. Completeness was agreed to the Jira change report and the production "
                "deployment log.",
                "Evidence 3 (Production_Deployment_Log.xlsx)",
                False,
            ),
            (
                procedures[2],
                "For the sampled change (CR-2025-0147), a Jira change request was confirmed to have "
                "been raised on 3 March 2025 with a technical specification attached, prior to development "
                "and subsequent deployment to production.",
                "Evidence 1 (Jira_CR_2025-0147.png)",
                False,
            ),
            (
                procedures[3],
                "UAT testing for CR-2025-0147 was confirmed to have been completed in the UAT environment "
                "by business owner C. Williams, with sign-off documented on 17 March 2025 — prior to "
                "production deployment on 20 March 2025.",
                "Evidence 2 (UAT_Signoff_CR-2025-0147.pdf)",
                False,
            ),
            (
                procedures[4],
                "CAB approval for CR-2025-0147 was confirmed within the Jira ticket, approved on "
                "10 March 2025, prior to UAT and production deployment. The approval record identifies "
                "the approving CAB members.",
                "Evidence 1 (Jira_CR_2025-0147.png)",
                False,
            ),
            (
                procedures[5],
                "The production deployment log confirms that CR-2025-0147 was deployed by Release Manager "
                "D. Chen on 20 March 2025. Developer B. Patel does not have production deployment access, "
                "confirming segregation of duties between the developer and the release function.",
                "Evidence 3 (Production_Deployment_Log.xlsx)",
                False,
            ),
            (
                procedures[6],
                "No emergency changes were identified in the sampled population. Management confirmed that "
                "1 emergency change was raised during FY2025 (CR-2025-0089); retrospective CAB approval "
                "was documented within 3 business days as required by policy. This was not included in the "
                "sample but was confirmed by inspection of the Jira ticket.",
                "Not evidenced for sample (no emergency change in sample)",
                False,
            ),
        ])
    ]

    ctrl.deficiencies = []

    ctrl.conclusion = (
        "Based on the evidence reviewed, the change management (code) control for Oracle EBS is "
        "effectively designed and implemented in accordance with the stated process. Code changes are "
        "formally requested via Jira, subject to CAB approval, tested in a UAT environment with "
        "business owner sign-off, and deployed to production exclusively by the Release Manager, "
        "ensuring appropriate segregation of duties. No control deficiencies were identified."
    )

    return ctrl


IHG_OUTPUT = "sample_output/Client_A_CyberArk_PAM_ITGC_Workpaper_FY2025_SAMPLE.xlsx"


def _make_privileged_access_cyberark() -> ITGCControl:
    ctrl = ITGCControl(
        control_type=ControlType.PRIVILEGED_ACCESS,
        application="CyberArk PAM",
        client_name="Client A",
        audit_period="FY2025",
        process_description=(
            "All privileged access to the client's production infrastructure — including servers, databases, "
            "and network devices — is managed through CyberArk Privileged Access Management (PAM). "
            "Privileged account credentials are stored in the CyberArk Enterprise Password Vault (EPV) "
            "and are never disclosed to end users in plain text. When a member of IT Operations or the "
            "Security team requires privileged access, they submit a request via the CyberArk web portal. "
            "The request is reviewed and approved by the IT Security Manager before the user is added to "
            "the relevant CyberArk Safe by the CyberArk Administrator. All privileged sessions are "
            "initiated exclusively through CyberArk Privileged Session Manager (PSM), which records and "
            "stores the full session for audit review. Credentials are automatically rotated by the "
            "Central Policy Manager (CPM) following each session checkout or on a maximum 30-day cycle. "
            "A quarterly review of all CyberArk Safe memberships is conducted by the IT Security team "
            "and signed off by the CISO."
        ),
    )

    procedures = TAILORED_PROCEDURES[ControlType.PRIVILEGED_ACCESS]

    ctrl.evidence_items = [
        EvidenceItem(
            filename="CyberArk_Safe_Members_Q4_2024.xlsx",
            file_path="",
            file_type="excel",
            order=1,
            hint="CyberArk Safe membership extract — Q4 2024 quarterly review",
            description=(
                "The spreadsheet extract from the CyberArk Enterprise Password Vault lists all Safe "
                "memberships as at 31 December 2024. The report shows 26 active privileged accounts "
                "across 9 Safes (Windows Servers, Oracle DB, Network Devices, CyberArk Admin, and "
                "5 application-specific Safes). Each entry records the account name, Safe name, "
                "Safe owner, and last password rotation date. Three accounts — svc_CTR_jenkins, "
                "svc_CTR_deploy, and A.Patel_admin — are annotated as belonging to contractors whose "
                "engagements ended between August and October 2024."
            ),
            key_findings=[
                "26 active privileged accounts across 9 CyberArk Safes documented.",
                "Safe structure separates accounts by platform type, supporting least privilege.",
                "3 contractor accounts (svc_CTR_jenkins, svc_CTR_deploy, A.Patel_admin) remain active "
                "despite contractor engagements ending in Q3 2024.",
                "Q4 2024 quarterly review sign-off present (CISO, dated 10 January 2025), but the 3 "
                "contractor accounts were not flagged for removal.",
            ],
            aligns_with_process="Partial",
            alignment_explanation=(
                "Safe membership report and quarterly review are in place; however, three contractor "
                "accounts were not identified for removal during the Q4 review, indicating the "
                "recertification process did not detect dormant privileged accounts."
            ),
            concerns=(
                "Three contractor privileged accounts remain active in the CyberArk vault beyond the "
                "end of the contractors' engagements (up to 5 months overdue for removal). These were "
                "not identified or actioned during the Q4 2024 quarterly recertification."
            ),
        ),
        EvidenceItem(
            filename="CyberArk_Access_Request_INC0091234.png",
            file_path="",
            file_type="image",
            order=2,
            hint="CyberArk access request and dual approval — sampled user R. Thompson",
            description=(
                "The screenshot depicts the CyberArk access request workflow for INC0091234, submitted "
                "on 6 January 2025 by R. Thompson (Senior Infrastructure Engineer). The request is for "
                "membership of the 'Windows_Prod_Servers' Safe. The workflow shows documented approval "
                "from the IT Security Manager (M. Okonkwo, 7 January 2025) prior to the CyberArk "
                "Administrator provisioning the Safe membership on 8 January 2025. A business "
                "justification referencing Project Phoenix server migration is attached."
            ),
            key_findings=[
                "Formal CyberArk access request raised via portal with documented business justification.",
                "IT Security Manager approval obtained (7 Jan 2025) prior to provisioning (8 Jan 2025).",
                "Access scoped to a single named Safe, consistent with least privilege.",
                "Request-to-provision turnaround: 2 business days, within the defined SLA.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "Access request, manager approval, and provisioning sequence align with the stated "
                "CyberArk request process."
            ),
            concerns="None noted",
        ),
        EvidenceItem(
            filename="CyberArk_PSM_Session_Log_Jan2025.xlsx",
            file_path="",
            file_type="excel",
            order=3,
            hint="PSM session recording log extract — January 2025",
            description=(
                "The spreadsheet extract from the CyberArk Privileged Session Manager (PSM) shows all "
                "privileged sessions initiated during January 2025. The log records 148 sessions, each "
                "with the initiating user, target system, Safe name, session start/end time, and a "
                "link to the recorded session file stored in the CyberArk Vault. No direct (non-PSM) "
                "connections to production systems are present in the firewall exception log cross-"
                "referenced in the rightmost column."
            ),
            key_findings=[
                "148 privileged sessions recorded in January 2025; all initiated via PSM.",
                "Session recordings linked for all entries — full audit trail maintained.",
                "No direct SSH/RDP connections to production systems detected outside PSM.",
                "All sessions attributable to named accounts — no shared credential usage observed.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "PSM session recording is operating as designed; all privileged access routed through "
                "PSM with no direct connections to production."
            ),
            concerns="None noted",
        ),
        EvidenceItem(
            filename="CyberArk_CPM_Password_Rotation.png",
            file_path="",
            file_type="image",
            order=4,
            hint="CyberArk CPM dashboard showing password rotation status",
            description=(
                "The screenshot depicts the CyberArk Central Policy Manager (CPM) dashboard as at "
                "31 January 2025. The dashboard shows 26 managed accounts, all displaying a 'Compliant' "
                "rotation status. The 'Last Changed' column confirms all credentials were rotated within "
                "the preceding 30 days. No accounts are flagged as 'Failed' or 'Pending Rotation'."
            ),
            key_findings=[
                "All 26 managed accounts show 'Compliant' rotation status.",
                "Most recent rotation dates fall within the 30-day policy maximum.",
                "Zero failed or pending rotation accounts at the point of inspection.",
                "CPM operating without manual intervention — automatic rotation confirmed.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "Automatic credential rotation by CPM is functioning for all vaulted accounts within "
                "the defined 30-day cycle."
            ),
            concerns="None noted",
        ),
    ]

    ctrl.procedure_results = [
        ProcedureResult(
            number=i + 1,
            procedure_text=proc,
            response=response,
            evidence_references=ev_ref,
            has_gap=gap,
        )
        for i, (proc, response, ev_ref, gap) in enumerate([
            (
                procedures[0],
                "The CyberArk Safe membership report (Q4 2024) was obtained and confirmed to represent "
                "the complete population of privileged accounts as at 31 December 2024, covering 26 "
                "accounts across 9 Safes. Completeness was agreed to the CyberArk EPV account inventory "
                "report provided by the CyberArk Administrator.",
                "Evidence 1 (CyberArk_Safe_Members_Q4_2024.xlsx)",
                False,
            ),
            (
                procedures[1],
                "For the sampled privileged user (R. Thompson, INC0091234), a formal CyberArk access "
                "request was confirmed, including a documented business justification (Project Phoenix "
                "server migration) and IT Security Manager approval dated 7 January 2025, prior to the "
                "Safe membership being provisioned on 8 January 2025.",
                "Evidence 2 (CyberArk_Access_Request_INC0091234.png)",
                False,
            ),
            (
                procedures[2],
                "The CyberArk Safe structure segregates accounts by platform (Windows Servers, Oracle DB, "
                "Network Devices, etc.), and the sampled user's access (Windows_Prod_Servers Safe only) "
                "is consistent with their role as Senior Infrastructure Engineer. No over-provisioning "
                "was identified for the sampled account.",
                "Evidence 1, Evidence 2",
                False,
            ),
            (
                procedures[3],
                "The CyberArk PSM session log confirms that all 148 privileged sessions during January "
                "2025 were initiated through PSM, with no direct connections to production systems "
                "detected. This architecture prevents the use of vaulted credentials for non-PSM "
                "sessions, enforcing the separation of privileged and day-to-day activity at the "
                "platform level.",
                "Evidence 3 (CyberArk_PSM_Session_Log_Jan2025.xlsx)",
                False,
            ),
            (
                procedures[4],
                "The PSM session log confirms full session recordings are captured for all privileged "
                "sessions, with recording links retained in the CyberArk Vault. The CPM dashboard "
                "confirms automatic credential rotation is active for all 26 managed accounts, "
                "providing a technical log of privileged account activity.",
                "Evidence 3, Evidence 4",
                False,
            ),
            (
                procedures[5],
                "All 26 privileged accounts in the CyberArk EPV are individually named accounts — no "
                "shared or generic privileged credentials were identified in the Safe membership report. "
                "The PSM session log confirms each session is attributable to a named user account.",
                "Evidence 1, Evidence 3",
                False,
            ),
            (
                procedures[6],
                "A quarterly CyberArk Safe membership recertification is conducted by IT Security and "
                "signed off by the CISO. However, the Q4 2024 review (signed off 10 January 2025) did "
                "not identify or action the removal of three contractor accounts (svc_CTR_jenkins, "
                "svc_CTR_deploy, A.Patel_admin) whose engagements ended between August and October 2024, "
                "indicating a gap in the effectiveness of the recertification process.",
                "Evidence 1 (CyberArk_Safe_Members_Q4_2024.xlsx)",
                True,
            ),
        ])
    ]

    ctrl.deficiencies = [
        ControlDeficiency(
            cd_number=1,
            procedure_number=7,
            description=(
                "The Q4 2024 CyberArk Safe membership recertification, signed off by the CISO on "
                "10 January 2025, failed to identify three privileged accounts belonging to contractors "
                "whose engagements with Client A had ended between August and October 2024 "
                "(svc_CTR_jenkins, svc_CTR_deploy, A.Patel_admin). These accounts remain active and "
                "vaulted in the CyberArk EPV up to five months after the contractors' departures, "
                "representing a risk of unauthorised privileged access to client production systems. "
                "The stated process requires the quarterly review to identify and remove accounts "
                "that are no longer required; the evidence indicates this did not occur for the "
                "above accounts."
            ),
            severity="Control Deficiency",
            management_response="Pending management response",
        )
    ]

    ctrl.conclusion = (
        "Based on the evidence reviewed, the privileged access management control for Client A "
        "Resorts, operated through CyberArk PAM, is substantially designed and implemented in "
        "accordance with the stated process. The CyberArk architecture — comprising EPV vaulting, "
        "PSM session recording, and CPM automatic credential rotation — provides a strong technical "
        "framework for managing and monitoring privileged access. Access requests are formally "
        "approved prior to provisioning, all sessions are recorded through PSM, and credentials "
        "are rotated within the defined 30-day cycle. However, one control deficiency (CD/W-1) "
        "has been raised as the Q4 2024 quarterly Safe membership recertification failed to identify "
        "and action the removal of three dormant contractor privileged accounts, indicating a gap "
        "in the operating effectiveness of the recertification process."
    )

    return ctrl


def _make_change_management_config_ihg() -> ITGCControl:
    ctrl = ITGCControl(
        control_type=ControlType.CHANGE_MANAGEMENT_CONFIG,
        application="CyberArk PAM",
        client_name="Client A",
        audit_period="FY2025",
        process_description=(
            "Configuration changes to CyberArk PAM and the client's production infrastructure are managed "
            "through the the client ServiceNow Change Management platform. All changes must have a Change "
            "Request (CR) raised in ServiceNow, including a technical specification, risk assessment, "
            "and a tested rollback plan. Standard changes — pre-approved routine tasks on the approved "
            "standard change catalogue — may be implemented without CAB approval but require a "
            "post-implementation review within 24 hours. Normal changes require approval from the "
            "bi-weekly Change Advisory Board (CAB), comprising IT Security, Infrastructure, and "
            "business stakeholders. Emergency changes may be approved by two members of the IT "
            "Security leadership team outside the CAB cycle, but retrospective CAB approval must be "
            "documented within 3 business days. All production configuration changes are implemented "
            "within approved change windows (Saturday 22:00–02:00 or Sunday 14:00–18:00 UTC) by the "
            "Infrastructure team. Members of the Security Engineering team who design changes "
            "do not have the ability to implement those changes in production."
        ),
    )

    procedures = TAILORED_PROCEDURES[ControlType.CHANGE_MANAGEMENT_CONFIG]

    ctrl.evidence_items = [
        EvidenceItem(
            filename="ServiceNow_CR_CHG0031847.png",
            file_path="",
            file_type="image",
            order=1,
            hint="ServiceNow change request for CyberArk platform upgrade to v14.2",
            description=(
                "The screenshot depicts ServiceNow Change Request CHG0031847, raised on 14 January 2025 "
                "by Security Engineer P. Nakamura for a CyberArk PAM platform upgrade from v14.0 to "
                "v14.2. The CR is classified as a Normal change and includes an attached technical "
                "specification, a risk assessment rated Medium, and a documented rollback plan. "
                "The CR status shows 'CAB Approved' with an approved change window of Saturday "
                "25 January 2025, 22:00–01:00 UTC."
            ),
            key_findings=[
                "Formal ServiceNow CR raised with technical specification, risk assessment, and rollback plan.",
                "Change correctly classified as Normal (not Standard), requiring CAB approval.",
                "CAB approval obtained prior to implementation window.",
                "Approved change window (Sat 25 Jan 2025) aligns with the defined production change schedule.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "CR documentation, risk assessment, rollback plan, and CAB approval all present prior "
                "to implementation, consistent with the Normal change process."
            ),
            concerns="None noted",
        ),
        EvidenceItem(
            filename="CAB_Minutes_22Jan2025.pdf",
            file_path="",
            file_type="pdf",
            order=2,
            hint="CAB meeting minutes — 22 January 2025",
            description=(
                "The PDF document contains the signed minutes of the Change Advisory Board meeting "
                "held on 22 January 2025. Agenda item 3 records the review and approval of CHG0031847 "
                "(CyberArk v14.2 upgrade). The minutes note that the CAB reviewed the technical "
                "specification and rollback plan, and confirmed the risk rating as appropriate. "
                "Approval was granted by all four CAB members present. The minutes are signed by the "
                "Change Manager (L. Adeyemi) and timestamped 22 January 2025 at 14:37 UTC."
            ),
            key_findings=[
                "CAB met on 22 January 2025 — 3 days before the approved change window.",
                "CHG0031847 reviewed and approved by all 4 CAB members present.",
                "Technical specification and rollback plan confirmed as reviewed by CAB.",
                "Minutes signed by Change Manager — formal record of approval maintained.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "CAB approval is formally documented in signed minutes prior to the implementation window."
            ),
            concerns="None noted",
        ),
        EvidenceItem(
            filename="Post_Implementation_Review_CHG0031847.xlsx",
            file_path="",
            file_type="excel",
            order=3,
            hint="Post-implementation review checklist for CHG0031847",
            description=(
                "The spreadsheet is the completed post-implementation review (PIR) checklist for "
                "CHG0031847, submitted on 26 January 2025 by Infrastructure Engineer T. Osei "
                "(the implementing engineer, distinct from the requesting Security Engineer). "
                "The checklist confirms: CyberArk upgraded to v14.2 successfully; all PSM session "
                "recordings tested and functional; CPM rotation verified post-upgrade; no service "
                "degradation observed; change implemented within the approved window "
                "(22:14–00:47 UTC). The PIR is countersigned by the Change Manager."
            ),
            key_findings=[
                "Change implemented by T. Osei (Infrastructure), not the requesting engineer P. Nakamura — SoD confirmed.",
                "CyberArk v14.2 upgrade completed within approved window (22:14–00:47 UTC, Sat 25 Jan).",
                "Post-implementation functional checks confirmed PSM and CPM operating normally.",
                "PIR submitted within 24 hours of implementation and countersigned by Change Manager.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "Implementation performed by a separate engineer from the requester, within the approved "
                "window, with a completed and signed PIR submitted within 24 hours."
            ),
            concerns="None noted",
        ),
        EvidenceItem(
            filename="Emergency_Change_CHG0031902.png",
            file_path="",
            file_type="image",
            order=4,
            hint="Emergency change for CyberArk CPM connectivity patch — out-of-hours",
            description=(
                "The screenshot depicts ServiceNow Emergency Change CHG0031902, raised on 3 February "
                "2025 at 02:17 UTC by Infrastructure Engineer T. Osei in response to a CPM connectivity "
                "failure affecting automated password rotation. The CR shows dual approval from the "
                "IT Security Director (S. Mensah, 02:23 UTC) and the Change Manager (L. Adeyemi, "
                "02:31 UTC) prior to implementation at 02:45 UTC. The retrospective CAB approval is "
                "recorded in the CR, referencing CAB minutes dated 5 February 2025 — 2 business days "
                "after the emergency change, within the 3-business-day requirement."
            ),
            key_findings=[
                "Emergency change raised with dual approval from IT Security Director and Change Manager prior to implementation.",
                "Implementation commenced at 02:45 UTC — 14 minutes after second approval received.",
                "Retrospective CAB approval documented within 2 business days (policy requires 3).",
                "Emergency change process followed in accordance with the stated policy.",
            ],
            aligns_with_process="Yes",
            alignment_explanation=(
                "Emergency change dual approval, implementation sequence, and retrospective CAB "
                "approval all comply with the defined emergency change policy."
            ),
            concerns="None noted",
        ),
    ]

    ctrl.procedure_results = [
        ProcedureResult(
            number=i + 1,
            procedure_text=proc,
            response=response,
            evidence_references=ev_ref,
            has_gap=gap,
        )
        for i, (proc, response, ev_ref, gap) in enumerate([
            (
                procedures[0],
                "The Change Management Policy (v4.1, approved March 2024) was obtained and "
                "inspected. The policy covers configuration changes to production infrastructure, "
                "including CyberArk PAM, and defines the Standard, Normal, and Emergency change "
                "classifications, approval requirements, and permitted change windows.",
                "Not evidenced (policy obtained separately from management)",
                False,
            ),
            (
                procedures[1],
                "Management provided the population of 89 configuration changes implemented during "
                "FY2025, extracted from ServiceNow. Completeness was agreed to the ServiceNow change "
                "report filtered for client production infrastructure CRs with status 'Closed — "
                "Successful' or 'Closed — Unsuccessful'.",
                "Not evidenced (population obtained from management via ServiceNow export)",
                False,
            ),
            (
                procedures[2],
                "For the sampled change (CHG0031847 — CyberArk v14.2 upgrade), a formal ServiceNow "
                "CR was confirmed to have been raised on 14 January 2025, including a technical "
                "specification, risk assessment (rated Medium), and rollback plan, prior to the "
                "approved implementation window on 25 January 2025. For the emergency change "
                "(CHG0031902), a CR was raised and dual approval obtained before implementation "
                "commenced.",
                "Evidence 1 (ServiceNow_CR_CHG0031847.png), Evidence 4 (Emergency_Change_CHG0031902.png)",
                False,
            ),
            (
                procedures[3],
                "CAB approval for CHG0031847 was documented in signed meeting minutes dated "
                "22 January 2025 — 3 days prior to the approved change window. All four CAB members "
                "present approved the change following review of the technical specification and "
                "rollback plan. For CHG0031902 (emergency change), retrospective CAB approval was "
                "obtained within 2 business days, within the 3-business-day policy requirement.",
                "Evidence 2 (CAB_Minutes_22Jan2025.pdf), Evidence 4 (Emergency_Change_CHG0031902.png)",
                False,
            ),
            (
                procedures[4],
                "The post-implementation review for CHG0031847 confirms the CyberArk v14.2 upgrade "
                "was tested following implementation, including functional verification of PSM session "
                "recording and CPM password rotation. No pre-implementation testing in a non-production "
                "environment is explicitly evidenced in the CR; management confirmed verbally that a "
                "staging environment was used, but no documented test results were provided.",
                "Evidence 3 (Post_Implementation_Review_CHG0031847.xlsx)",
                False,
            ),
            (
                procedures[5],
                "The the client ServiceNow change log and CPM connectivity alert log are reviewed by the "
                "Change Manager on a weekly basis, as confirmed by management. Evidence of the weekly "
                "log review was not provided as part of this sample; management indicated this review "
                "is performed informally without a documented sign-off.",
                "Not evidenced (management confirmation only)",
                False,
            ),
            (
                procedures[6],
                "The post-implementation review checklist for CHG0031847 includes a documented rollback "
                "procedure; the PIR confirms the change was successful and the rollback was not "
                "required. For CHG0031902, the emergency change CR includes a rollback step. Management "
                "confirmed rollback procedures are tested annually during DR exercises, although "
                "specific DR test evidence was not reviewed as part of this sample.",
                "Evidence 3 (Post_Implementation_Review_CHG0031847.xlsx)",
                False,
            ),
        ])
    ]

    ctrl.deficiencies = []

    ctrl.conclusion = (
        "Based on the evidence reviewed, the change management (configuration) control for CyberArk "
        "PAM and Client A' production infrastructure is effectively designed and "
        "implemented in accordance with the stated process. Configuration changes are formally "
        "requested in ServiceNow with appropriate documentation, subject to CAB approval prior to "
        "implementation, deployed within approved change windows by engineers separate from the "
        "requesting team (confirming segregation of duties), and supported by completed post-"
        "implementation reviews. The emergency change process was also observed to operate in "
        "accordance with policy, with dual approval and timely retrospective CAB sign-off. "
        "No control deficiencies were identified."
    )

    return ctrl


if __name__ == "__main__":
    import os
    os.makedirs("sample_output", exist_ok=True)

    # Sample 1: Generic Oracle EBS (original)
    controls_oracle = [
        _make_access_provisioning(),
        _make_change_management(),
    ]
    generate_workpaper(controls_oracle, OUTPUT)
    print(f"Generated: {OUTPUT}")

    # Sample 2: Client A — CyberArk PAM focus
    controls_ihg = [
        _make_privileged_access_cyberark(),
        _make_change_management_config_ihg(),
    ]
    generate_workpaper(controls_ihg, IHG_OUTPUT)
    print(f"Generated: {IHG_OUTPUT}")
