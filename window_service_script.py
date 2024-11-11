import win32serviceutil
import win32service
import win32event
from report import server_start

class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ITReport"
    _svc_display_name_ = "IT Report Service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        # import servicemanager
        # servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
        #                       servicemanager.PYS_SERVICE_STARTED, (self._svc_name_, ""))
        server_start()


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(MyService)
