import email
import os
import re
import csv


path = "email_files"
listing = os.listdir(path)


def append_to_csv(row_message):
    writer = None
    csvfile = open("registered_voters.csv", "a")
    fieldnames = ["FullName", "Email", "IdNumber", "ContactNumber", "Location"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writerow(tuple_to_dict(zip(fieldnames, row_message), {}))


def tuple_to_dict(tup, di):
    di = dict(tup)
    return di


for fle in listing:
    if str.lower(fle[-3:]) == "eml":
        msg = email.message_from_file(open("./{0}/{1}".format(path, fle)))
        attachments = msg.get_payload()

        try:
            message = attachments[0].get_payload(decode=True)
            message = re.sub("[\s+]", "", message)
            row_message = [col.split(":")[1] for col in message.split(",")]
            append_to_csv(row_message)
            print("Logger: Successfully added row in csv file!")
        except Exception as detail:
            print(detail)
            pass
