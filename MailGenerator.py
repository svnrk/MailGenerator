
class Email:

    def __init__(self, address, mail_topic, content, attachment):
        self.address = address
        self.mail_topic = mail_topic
        self.content = content
        self.attachment = attachment

    def __str__(self):
        return ("Address: %s, Topic: %s" % (self.address, self.mail_topic))

def mailAddressFromExcel(excel_path):
    # https: // www.geeksforgeeks.org / reading - excel - file - using - python /
    pass

def mailTopicSetter():
    pass

def mailContentFromFile(content_path):
    pass

def mailAttachmentFromFolder(folder_path, document_name):
    pass

def mailComposer():
    pass

def mailSender():
    pass


next_email = Email("xxx@xxx.xxx", "test", "testtesttest", "false")

print (next_email.attachment)
