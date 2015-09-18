import main
import constants
p=main.paperwork()
p.generate_utp()
p.generate_code_review()
p.generate_STR_form()
m=main.mailer(p)
m.chrq_approval_mail()
m.patch_mail()
m.com_request_mail()