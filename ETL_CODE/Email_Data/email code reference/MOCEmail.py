from dataclasses import dataclass
from body_parsing import MOCEmailNotificationBody
from datetime import date


@dataclass
class MOCEmail:
    sender: str
    body: MOCEmailNotificationBody
    subject: str
    sent_on: date

    @staticmethod
    def from_outlook_message(message):
        return MOCEmail(
            sender=str(message.Sender),
            body=MOCEmailNotificationBody.parse_body(message.body),
            subject=str(message.Subject),
            sent_on=message.senton.date().isoformat()
        )
