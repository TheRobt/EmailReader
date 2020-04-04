import win32com.client
import os

class EmailRead():
	def __init__(self, email, box):
		self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
		self.messages = self.outlook.Folders(email).Folders(box).Items
		try:
			os.mkdir("c:\Attachments")
		except (FileExistsError):
			pass
	def __tryRead(self,message,comand):
		Dict_from_data = {
			"Subject":message.Subject,  
			"ReceivedTime":message.ReceivedTime,
			"EntryID":message.EntryID,
			"HtmlBody":message.HtmlBody,  
			"Size":message.Size,
			"SenderName":message.SenderName,
			"To":message.To,			      
			"Cc":message.Cc,		      
			"Body":message.Body
			}
		try:
			Dict_from_data[comand]
		except:
			return None
		return Dict_from_data[comand]
	def __getMessage(self, message, With_Attachments=None):
		email_read ={"Attachments"	:[]}
		try:
			email_read["Subject"]=message.Subject
		except:
			email_read["Subject"]=None
		try:
			email_read["ReceivedTime"]=message.ReceivedTime
		except:
			email_read["ReceivedTime"]=None
		try:
			email_read["EntryID"]=message.EntryID
		except:
			email_read["EntryID"]=None
		try:
			email_read["HtmlBody"]=message.HtmlBody
		except:
			email_read["HtmlBody"]=None
		try:
			email_read["Size"]=message.Size
		except:
			email_read["Size"]=None
		try:
			email_read["SenderName"]=message.SenderName
		except:
			email_read["SenderName"]=None
		try:
			email_read["To"]=message.To
		except:
			email_read["To"]=None
		try:
			email_read["Cc"]=message.Cc
		except:
			email_read["Cc"]=None
		try:
			email_read["Body"]=message.Body
		except:
			email_read["Body"]=None
		if((message.Attachments) and(With_Attachments)):	
			try:
				os.mkdir(str("c:\Attachments\\"+email_read['EntryID']))
			except (FileExistsError):
				pass
			self.path = str("c:\Attachments\\"+email_read['EntryID'])
			for attachment in message.Attachments:
					attachment.SaveAsFile(os.path.join(self.path, str(attachment)))
					attachments_tuple = (str(attachment),os.path.join(self.path, str(attachment)))
					email_read["Attachments"].append(attachments_tuple)
		return email_read
	def len_of_box(self):
		return len(self.messages)
	def get_first_message(self, With_Attachments=None):
		return self.__getMessage(self.messages.GetFirst(), With_Attachments)
	def get_last_message(self, With_Attachments=None):
		return self.__getMessage(self.messages.GetLast(), With_Attachments)
	def get_next_message(self, With_Attachments=None):
		message=self.messages.GetNext()
		if(message):
			return self.__getMessage(message, With_Attachments)
		else:
			return None
	def get_previous_message(self,With_Attachments=None):
		message=self.messages.GetPrevious()
		if(message):
			return self.__getMessage(message, With_Attachments)
		else:
			return None
	def get_all_messages(self,With_Attachments=None):
		message=self.messages.GetFirst()
		Messages={}
		i = 0
		while message:
			#cccccccint(message)
			i+=1
			try:
				Messages[str(i)]=self.__getMessage(message, With_Attachments)
			except:
				pass
			message=self.messages.GetNext()
		return Messages
	def get_message_by_number(self, number,With_Attachments=None):
		return self.__getMessage(self.messages.Item(number),With_Attachments)

email = EmailRead('lucas.osantana@viaquatro.com.br',"Caixa de Entrada")
email.get_all_messages()
#email.get_previous_message()
