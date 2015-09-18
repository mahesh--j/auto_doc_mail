#Template base documents to which we add appropriate details
utp_base_document="C:\\base_documents\\unit_test_plan_base.xlsx"
code_review_base_document="C:\\base_documents\\code_review_base.xlsx"
str_base_document="C:\\base_documents\\str_base.docx"

#Directory Path
utp_path="C:\\docs\\test\\"
str_path="C:\\docs\\test\\"
code_review_path="C:\\docs\\test\\"

#Development details
initials="MJG"
trga="MJG"
developer_name="Mahesh Jagtap"
email="mahesh.jagtap@sungard.com"
extension="2966"
authoriser="Abhay Kenjalkar"
development_manager="Mitesh Patel"
client_name="HSBC"

#Environment details
test_environment="DEV115"
RIMS_version = "RLS 11.5"
temp_dir="C:\\temp\\"

#STR form parameters
billable_qa = "N"
system_name = "SYS 11.5"
system_version = "11.5"
days_for_delivery_date = 4 	#Expected delivery date to QA = SYSDATE + days_for_delivery_date
days_for_qa_testing = 2		#Date requested to be tested by = Expected delivery date to QA + days_for_qa_testing
complexity = "M"
priority = "H"
special_note = "Due to the bespoke nature of the change , we are requesting dispensation for the system test but will provide details of the testing conducted in the release area for QA validation"

#CHRQ Approval mail
chrq_apprival_mail_reqd = "Y"
chrq_approval_greet = "Hi Abhay,"
chrq_approval_to = 'Abhay.Kenjalkar@sungard.com'
chrq_approval_cc = 'Amey.Chaudhari@sungard.com'

#Patch mail 
patch_mail_reqd = "Y"
patch_greet = "Hello Team,"
patch_to = 'CM.PTS.Support.DBAClient@sungard.com'
patch_cc = 'CM.PTS.Implementation.HSBC@sungard.com'

#COM request mail
com_request_mail_reqd = "Y"
com_request_mail_greet = "Hello Team"
com_request_mail_to = 'Ahmed.Ouni@sungard.com; Sunil.Ramidi@sungard.com; Abir.GUEDDICHE@sungard.com; CM.PTS.QA.Managers@sungard.com'
com_request_mail_cc = 'CM.PTS.Implementation.HSBC@sungard.com'
