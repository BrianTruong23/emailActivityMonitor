import pandas as pd
from app import db, Message, app
from app import db, app

with app.app_context():
    db.create_all()

df = pd.read_excel("result/log2.xlsx")

with app.app_context():
    for _, row in df.iterrows():
        msg = Message(
            message_id=row["Message_ID"],
            email_sender=row["Email Sender"],
            date=row["Date"],
            time_received=row["Time Received"],
            wait_time=row["Wait Time"],
            status=row["Status"]
        )
        db.session.add(msg)
    db.session.commit()
