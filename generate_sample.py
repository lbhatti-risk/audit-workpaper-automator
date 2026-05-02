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


if __name__ == "__main__":
    import os
    os.makedirs("sample_output", exist_ok=True)

    controls = [
        _make_access_provisioning(),
        _make_change_management(),
    ]

    generate_workpaper(controls, OUTPUT)
    print(f"Sample workpaper generated: {OUTPUT}")
