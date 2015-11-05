import win32com.client
import sys
import pywintypes
import logging

#Get an instance of a logger.
logger = logging.getLogger()
#Get a handler to use in logging.
handler = logging.StreamHandler(sys.stdout)
#Set the log to display in the console.
logger.setLevel(logging.DEBUG)
#Add the custom handler.
logger.addHandler(handler)
#Set level to display the logs in the console.
logger.setLevel(logging.DEBUG)

#Initialize a constant to store the Quality Center App name.
CONST_TD_APP_NAME = "TDApiOle80.TDConnection"


class QC_Connection(object):
    """QC_Connection Class

    This class connects to Quality Center (QC) using the Open Test 
    Architecture (OTA) API. 

    """

    def __init__(self):
        logger.info ("init")
    
    def login_to_qc(self, server, username, password, domain, project):
        '''
        This method is used to create a Quality Center connection instance, 
        and login to it.
        '''
        try:
            logger.info("Attempting to connect to Quality Center using OTA API...")
            #Initialize the COM object to dispatch with respect to the OTA API. 
            qc = win32com.client.Dispatch(CONST_TD_APP_NAME)
            #Initialize the connection with the QC server using the dispatch object.
            qc.InitConnectionEx(server)
            #Attempt to login using the username and password.
            qc.login(username, password)
            #Once the login is successful, 
            #Connect to the appropriate DOMAIN and PROJECT.
            qc.Connect(domain, project)
            #Verify that the user is logged in.
            if qc.LoggedIn:
                logger.info("Logged in to Quality Center using the OTA API.")
            else:
                logger.exception("The user, %s could not login." % username)
                return None
            #Return the TD Connection object.
            return qc
        except pywintypes.com_error as e:
            #Store the error message in the COM exception.
            exception_message = e.excepinfo[2]
            #Log the error message found in the exception.
            logger.error(exception_message)
            return None