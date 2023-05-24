from RPA.Outlook.Application import Application
from datetime import datetime


class ExtendedOutlook(Application):
    def __init__(self, autoexit: bool = True) -> None:
        super().__init__(autoexit)

    def get_active_email(self):
        """New keyword for getting user's current active email item

        :return: dictionary of email content or None
        """
        explorer = self.app.ActiveExplorer()
        active_email = None
        if explorer.Selection.Count > 0:
            active_email = explorer.Selection.Item(1)
            return self._mail_item_to_dict(active_email)

    def _get_sender_email_address(self, mail_item):
        mi = mail_item
        try:
            return (
                mi.Sender.GetExchangeUser().PrimarySmtpAddress
                if mi.SenderEmailType == "EX"
                else mi.SenderEmailAddress
            )
        except AttributeError:
            return None

    def _mail_item_to_dict(self, mail_item):
        mi = mail_item
        response = {
            "Sender": self._get_sender_email_address(mi),
            "To": [],
            "CC": [],
            "BCC": [],
            "Subject": mi.Subject,
            "Body": mi.Body,
            "Attachments": [
                {"filename": a.FileName, "size": a.Size, "item": a}
                for a in mi.Attachments
            ],
            "Size": mi.Size,
            "object": mi,
        }
        rt = getattr(mail_item, "ReceivedTime", "<UNKNOWN>")
        response["ReceivedTime"] = rt.isoformat() if rt != "<UNKNOWN>" else rt
        response["ReceivedTimestamp"] = (
            datetime(
                rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second
            ).timestamp()
            if rt and not isinstance(rt, str)
            else None
        )
        so = getattr(mail_item, "SentOn", "<UNKNOWN>")
        response["SentOn"] = so.isoformat() if so != "<UNKNOWN>" else so
        if hasattr(mi, "Recipients"):
            self._handle_recipients(mi.Recipients, response)
        return response


if __name__ == "__main__":
    app = ExtendedOutlook(autoexit=False)
    app.open_application()
    active_email = app.get_active_email()
    print(f"SUBJECT = {active_email.Subject}")
    print(f"SENDER = {active_email.SenderName}")
