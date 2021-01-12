from os import system
from mailmerge import MailMerge
from datetime import date

def remittance_generator():
    template = "remittance_advice.docx"
    document = MailMerge(template)
    print("Merge fields : ",document.get_merge_fields())
    output_name = get_name()
    merge_date = get_date()
    formatted_payment = get_payment()
    reference = get_reference(input_name)
    print("NAME:", output_name)
    print("DATE: ", merge_date)
    print("PAYMENT: ", formatted_payment)
    print("REFERENCE: ", reference)
    #MERGE FIELDS IN DOC
    document.merge(name=output_name,date=merge_date,payment=formatted_payment)
    filename = "done.docx"
    document.write(filename)
    cmd = "open {0} -a 'Microsoft Word'".format(filename)
    system(cmd)

def get_name():
    global input_name
    input_name = input("Enter name: ")
    if input_name in bg_team_dict:
        output_name = bg_team_dict.get(input_name)
        print('BG team  :', output_name)
        return output_name
    output_name = input_name
    return output_name

def get_date():
    merge_date = input("Enter date: ")
    if len(merge_date) == 0:
        today = date.today()
        merge_date = today.strftime("%d/%m")
        return merge_date
    return merge_date

def get_payment():
    payment = int(input("Enter remittance amount: "))
    formatted_payment = '{0:.2f}'.format(payment)
    return formatted_payment

def get_reference(input_name):
    bg_team = input("BG team? y or blank ")
    if bg_team == "y":
        reference = bg_team_refs.get(input_name)+1
        return reference
    reference = ""
    return reference

if __name__ == '__main__':
    remittance_generator()
