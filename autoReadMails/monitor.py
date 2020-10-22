from msal import PublicClientApplication
import json,re,sys
import msal,requests 
import pandas as pd
from requests_oauthlib import OAuth2Session
import pybase64
import logging

logging.basicConfig(level = logging.DEBUG,
filename = "logs.log",
format = "%(levelname)s: %(message)s -%(asctime)s")

class MailboxMonitor:
	
	"""Class to Monitor a Mailbox

	Attributes:
	config(dict): Mailbox configuration obtained from microsoft
		graph application
	app(obj): configured msal application object

	Methods:
	get_access_token(): function to obtain access token 
	get_attachment_in_file(): function to download attachment for a given mail
	get_message_details(): function to get details of the given mail
	read_sent_mails(): function to read mails from sent items folder
	read_folder_mails(): function to read mails from a given folder
	read_inbox_mails(): function to read mails from inbox
	get_conversation_thread(): function to read the latest mail on a given 
		mail chain or conversation thread
	"""
	def __init__(self,config):
		self.config = json.load(open(config))
		self.app = msal.PublicClientApplication(
			self.config["client_id"], authority=self.config["authority"],
			)
			
	def get_access_token(self):

		"""Function to get the access token by authenticating
		using username and password

		parameters:
		None

		returns:
		None
		"""
		try:
			result = None
			accounts = self.app.get_accounts(username = self.config["username"])
			if accounts:
				for a in accounts:
					pass
				chosen = accounts[0]
				result = self.app.acquire_token_silent(["User.Read","Group.Read.All","Group.ReadWrite.All"], account=chosen)	
			
			if not result:
				result = self.app.acquire_token_by_username_password(
					self.config["username"], self.config["password"], scopes=["User.Read","Mail.ReadWrite.Shared","Mail.ReadWrite"])
				
			if "access_token" in result:
				self.access_token = result["access_token"]
				logging.info("Access token obtained")
				pass
			else:
				logging.error(result.get("error"))
				logging.error(result.get("error_description"))
				logging.error(result.get("correlation_id"))
				
		except Exception as ex:
			logging.error("Exception in get access token - "+str(ex))
			return -1
			
	def get_attachment_in_file(self,message_id:str,file:str)->None:

		"""Function to download the attachment in the given messsage id
		to the provided full file path

		parameters:
		message_id(str)* : Id of the message from which the 
			attachment is to be downloaded
		file(str)* : full path including filename to which the
			the attachment is to be saved

		returns:
		None
		"""
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/messages/{1}/attachments".format(self.config['endpoint'],message_id)
			response = requests.get(URL,headers=self.headers)
			print(response)
			data = json.loads(response.text)
			attachment_id = data["value"][0]["id"]
		except Exception as ex:
			logging.error("No attachment ID for message- "+str(message_id)+" - "+str(ex))
			return -1
			
		try:
			URL = "{0}/me/messages/{1}/attachments/{2}".format(self.config['endpoint'],message_id,attachment_id)
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			encoded_data = data["contentBytes"]
			file_data  = pybase64.b64decode(encoded_data, altchars=None, validate=False)
			with open(file,"wb") as fp:
				fp.write(file_data)
			logging.info("Attachment saved as : "+str(file))
			return 0
		except Exception as ex:
			logging.error("Ex. in get attachment in file "+str(ex))
			return -1

	def get_message_details(self,message_id:str)->"Message details(dict)":

		"""Function to fetch conversation ID,subject,body,ToRecipients,
		CC Recipients from message ID

		parameters:
		message_id(str)* : unique ID of the email 

		returns:
		dict_message_details(dict)* : Dictionary containing message details
		"""
		dict_message_details = {}
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/messages/{1}".format(self.config['endpoint'],message_id)
			response = requests.get(URL,headers=self.headers)
			print(response)
			data = json.loads(response.text)
			dict_message_details["conversationId"] = data["conversationId"]
			dict_message_details["subject"] = data["subject"]
			dict_message_details["body"] = data['bodyPreview']
			dict_message_details["from"] = data["from"]["emailAddress"]["address"]
			dict_message_details["toRecipients"] = [address["emailAddress"]["address"] for address in data["toRecipients"]]
			dict_message_details["ccRecipients"] = [address["emailAddress"]["address"] for address in data["ccRecipients"]]
			return dict_message_details
				
		except Exception as ex:
			logging.error("No details obtined for message- "+str(message_id)+" - "+str(ex))
			
	def read_sent_mails(self):
		
		"""fetches the last 5 emails from the sent
		items folder

		parameters:
		None

		returns:
		dict_log(dict): nested dictionary with message id as key
			and message details dictionary as value
		"""
		dict_log = {}
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/mailFolders?top=100".format(self.config['endpoint'])
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:
				if str(item['displayName']).strip().lower() == "sent items":
					print(item['id'])
					folder_id = item['id']
					break
			URL = "{0}/me/mailFolders/{1}/messages?top=5".format(self.config['endpoint'],folder_id)
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:
				dict_log[item['id']] = self.get_message_details(item['id'])
			return dict_log
				
		except Exception as ex:
			logging.error("exception in get sent item - "+str(ex))

	def get_conversation_thread(self,conversation_id:str):

		"""Function to get the latest conversation on a mail chain

		parameters:
		conversation_id(str)*: conversation id of the mail chain

		returns:
		dict_conversation_details(dict) : dictionary containing details 
			of the latest update on the mail thread
		"""
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/mailFolders?top=100".format(self.config['endpoint'])
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:	
				if str(item['displayName']).strip().lower() == "inbox":
					inbox_id = item['id']
			URL = "{0}/me/mailFolders/{1}/messages?top=100".format(self.config['endpoint'],inbox_id)
			response = 	requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:	
				if str(item['conversationId']).strip() == str(conversation_id).strip():
					dict_conversation_details = self.get_message_details(item["id"])
					return dict_conversation_details
			return {}

		except Exception as ex:
			logging.error("Exception in get conversation thread - "+str(ex))
					
	def read_folder_mails(self,folder_name:str):
		
		"""Function to read the latest unread mail from a 
		given inbox folder
		
		parameters:
		folder_name(str)* : Folder from which the mail is to be read

		returns:
		message_details(dict) : dictionary containg latest message details
		"""
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/mailFolders?top=100".format(self.config['endpoint'])
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:	
				if str(item['displayName']).strip().lower() == "inbox":
					inbox_id = item['id']
					break
			URL = "{0}/me/mailFolders/{1}/childFolders?top=100".format(self.config['endpoint'],inbox_id)
			response = 	requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:
				if str(item['displayName']).strip().lower() == str(folder_name).lower():
					folder_id = item['id']
					break
			URL = "{0}/me/mailFolders/{1}/childFolders/{2}/messages?filter=isRead eq false&top=1".format(self.config['endpoint'],inbox_id,folder_id)
			response = 	requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			message_id = data['value'][0]['id']
			message_details = self.get_message_details(message_id)
			URL = "{0}/me/mailFolders/{1}/childFolders/{2}/messages/{3}".format(self.config['endpoint'],inbox_id,folder_id,message_id)
			self.headers_new = {
			"Authorization": "Bearer {}".format(self.access_token),
			"Content-Type" : "application/json",
			"accept":"application/json"
			}
			payload=json.dumps({"isRead": "true"})
			response = requests.patch(URL,headers=self.headers_new,data=payload)
			logging.info(str(response.text))
			return message_details	
		
		except Exception as ex:
			logging.error("Exception in get read folder mails - "+str(ex))
			
	def read_inbox_mails(self):

		"""function to read latest unread mail from inbox

		parameters:
		None

		returns:
		message_details(dict) : dictionary containing latest message details
		"""
		self.get_access_token()
		self.headers = {"Authorization": "Bearer {}".format(self.access_token)}
		try:
			URL = "{0}/me/mailFolders?top=100".format(self.config['endpoint'])
			response = requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			for item in data['value']:	
				if str(item['displayName']).strip().lower() == "inbox":
					inbox_id = item['id']
			URL = "{0}/me/mailFolders/{1}/messages?filter=isRead eq false&top=1".format(self.config['endpoint'],inbox_id)
			response = 	requests.get(URL,headers=self.headers)
			data = json.loads(response.text)
			message_id = data['value'][0]['id']
			message_details = self.get_message_details(message_id)
			URL = "{0}/me/mailFolders/{1}/childFolders/{2}/messages/{3}".format(self.config['endpoint'],inbox_id,folder_id,message_id)
			self.headers_new = {
			"Authorization": "Bearer {}".format(self.access_token),
			"Content-Type" : "application/json",
			"accept":"application/json"
			}
			payload=json.dumps({"isRead": "true"})
			response = requests.patch(URL,headers=self.headers_new,data=payload)
			logging.info(str(response.text))
			return message_details
		except Exception as ex:
			logging.error("Exception in read inbox mail - "+str(ex))

if __name__ == "__main__":
	my_mailbox = MailboxMonitor("config1.json")
	while True:
		try:
			recent_unread_mail = my_mailbox.read_folder_mails("test_folder")
			print(recent_unread_mail)
		except Exception as ex:
			logging.error(str(ex))
		
			
	
		