from typing import List
from pydantic import BaseModel

class EmailOptions(BaseModel):
    selected_profiles: List[str]
    browser_language: str
    send_limit_per_profile: int
    loop_profile: int

class EmailPayload(BaseModel):
    subject: str
    email_list: str
    body: str
    options: EmailOptions
