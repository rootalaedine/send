from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust this based on your requirements
    allow_methods=["*"],
    allow_headers=["*"],
)

class EmailPayload(BaseModel):
    subject: str
    email_list: str
    body: str

class EmailOptions(BaseModel):
    selected_profiles: str
    browser_language: str
    send_limit_per_profile: int
    loop_profile: int

@app.post("/send_email")
async def send_email(payload: EmailPayload, options: EmailOptions):
    # Print the received payload and options
    print("*"*50,"api.py","*"*50)
    
    print("Received Payload:", payload,"\n")
    
    print("Received Options:", options,"\n")
    
    print("*"*50,"api.py","*"*50)

    # Process the data and send the response

    # Return a response indicating success or failure
    return {"message": "Email sent successfully",
            "email payload" : payload,
            "options emails" :options
            }

