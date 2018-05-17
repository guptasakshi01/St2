'''
Python File to monitor mailbox and log the sender of email
'''
from exchangelib import Credentials, Account, DELEGATE, Configuration, Mailbox
from st2reactor.sensor.base import PollingSensor
import eventlet
from st2client.client import Client


__all__ = [
    'MonitorMailbox'
]

eventlet.monkey_patch(
    os=True,
    select=True,
    socket=True,
    thread=True,
    time=True)

class MonitorMailbox(PollingSensor):
    '''
    Sensor class to monitor mail box
    '''
    def __init__(self, sensor_service, config=None, poll_interval=1):
        '''
        Method toinitialize and fetch required details from datastore
        '''
        super(MonitorMailbox, self).__init__(sensor_service=sensor_service,\
                                            config=config, \
                                            poll_interval=poll_interval)
        self._trigger = 'hivecenter_msexchange.monitor.mailbox'
        self._logger = self._sensor_service.get_logger(__name__)
        client = Client(base_url='http://localhost')
        exchangeuser = client.keys.get_by_name(name="exchangeusername").value
        exchangepassword = client.keys.get_by_name(name="exchangepassword").value
        credentials = Credentials(username=exchangeuser, password=exchangepassword)
        config = Configuration(server='Outlook.Office365.com', credentials=credentials)
        self.account = Account(primary_smtp_address=exchangeuser, credentials=credentials, autodiscover=False, config=config, access_type=DELEGATE)

    def setup(self):
        '''
        Setup for sensor
        '''
        self._logger.debug('[MonitoringMailbox] : entering setup')

    def poll(self):
        '''
        This Method is CRUX of system, it is called on every poll interval
        '''
        self._logger.debug('[MonitoringMailbox]: entering poll')
        self.poll_mailbox()

    def cleanup(self):
        '''
        This is called when the st2 system goes down.
        '''
        self._logger.debug('[MonitoringMailbox]: entering cleanup')
        pass

    def add_trigger(self, trigger):
        '''
        This method is called when trigger is created
        '''
        pass

    def update_trigger(self, trigger):
        '''
        This method is called when trigger is updated
        '''
        pass

    def remove_trigger(self, trigger):
        '''
        This method is called when trigger is deleted
        '''
        pass

    def poll_mailbox(self):
        '''
        This method polls mailbox every one second
        '''
        self._logger.debug('[Entering Poll Mailbox]')
        '''
        Monitors unread mails containg subject test
        '''
        for item in self.account.inbox.filter(subject__contains='test', is_read=False):
            item.is_read = True
            item.save()
            self._logger.debug(item.author.email_address)

    def _process_message(self, payload, trace=None):
        '''
        Method to Dispatch Trigger
        '''
        self._sensor_service.dispatch(trigger=self._trigger, payload=payload)
