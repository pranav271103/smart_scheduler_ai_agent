import pyttsx3
import speech_recognition as sr
import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import dateparser
from dateutil import parser
from datetime import timedelta
import pytz
import os
import re
import sys
import json
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
from googletrans import Translator
import pickle

class SmartScheduler:
    def __init__(self):
        self.input_mode = None
        self.output_mode = None
        self.creds = None
        self.calendar_service = None
        self.meeting_details = {
            'duration': None,
            'time': None,
            'day': None,
            'title': None,
            'attendees': [],
            'reminders': [],
            'language': 'en'
        }
        self.user_preferences = self.load_user_preferences()
        self.engine = None
        self.recognizer = None
        self.llm = self.initialize_gemini()
        self.translator = Translator()
        self._current_conversation_context = []
        
        try:
            self.authenticate_calendar()
        except Exception as e:
            print(f"Authentication failed: {str(e)}")
            print("Please check your credentials and try again.")
            sys.exit(1)

    def initialize_gemini(self):
        genai.configure(api_key='YOUR API KEY')  # API KEY - Put Yours Here
        return genai.GenerativeModel('gemini-1.5-flash-latest') # let the model be the same as gemini-1.5-flash-latest because it supports this model as per latest rate limits.

    def load_user_preferences(self):
        try:
            if os.path.exists('user_preferences.pkl'):
                with open('user_preferences.pkl', 'rb') as f:
                    return pickle.load(f)
        except Exception as e:
            print(f"Error loading preferences: {e}")
        return {
            'usual_meeting_times': {},
            'preferred_language': 'en',
            'default_reminders': ['email', 10]
        }

    def save_user_preferences(self):
        try:
            with open('user_preferences.pkl', 'wb') as f:
                pickle.dump(self.user_preferences, f)
        except Exception as e:
            print(f"Error saving preferences: {e}")

    def translate_text(self, text, dest_language=None):
        if not dest_language:
            dest_language = self.meeting_details['language']
        try:
            translated = self.translator.translate(text, dest=dest_language)
            return translated.text
        except Exception as e:
            print(f"Translation error: {e}")
            return text

    def generate_conversational_response(self, user_input, context=None):
        safety_settings = {
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE
        }
        
        conversation_history = "\n".join([f"User: {msg['input']}\nAssistant: {msg['response']}" 
                                       for msg in self._current_conversation_context[-3:]])
        
        prompt = f"""You are SmartScheduler, a multilingual AI assistant for scheduling meetings.
Current language: {self.meeting_details['language']}
User preferences: {json.dumps(self.user_preferences, indent=2)}

Conversation history:
{conversation_history}

Current meeting details:
- Duration: {self.meeting_details['duration'] or "Not specified"}
- Time: {self.meeting_details['time'] or "Not specified"}
- Day: {self.meeting_details['day'] or "Not specified"}
- Title: {self.meeting_details['title'] or "Not specified"}
- Attendees: {', '.join(self.meeting_details['attendees']) if self.meeting_details['attendees'] else "None"}
- Reminders: {', '.join(str(r) for r in self.meeting_details['reminders']) if self.meeting_details['reminders'] else "None"}

{context or ''}

User: {user_input}

Respond conversationally and help complete the scheduling process. Ask relevant follow-up questions.
If suggesting times, consider the user's usual meeting times with these attendees."""
        
        try:
            response = self.llm.generate_content(
                prompt,
                safety_settings=safety_settings
            )
            return response.text
        except Exception as e:
            print(f"LLM error: {e}")
            return "I encountered an error. Let's try again."

    def check_for_conflicts(self):
        start_time = self.parse_time(f"{self.meeting_details['day']} {self.meeting_details['time']}")
        if not start_time:
            return None
        
        end_time = start_time + timedelta(minutes=self.meeting_details['duration'])
        
        try:
            events_result = self.calendar_service.events().list(
                calendarId='primary',
                timeMin=start_time.isoformat(),
                timeMax=end_time.isoformat(),
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            return events_result.get('items', [])
        except Exception as e:
            print(f"Error checking conflicts: {str(e)}")
            return None

    def suggest_usual_times(self, attendees):
        usual_times = {}
        for attendee in attendees:
            if attendee in self.user_preferences['usual_meeting_times']:
                usual_times.update(self.user_preferences['usual_meeting_times'][attendee])
        
        if usual_times:
            return max(usual_times.items(), key=lambda x: x[1])[0]
        return None

    def handle_conflict_resolution(self, conflicts):
        conflict_details = "\n".join([f"- {event['summary']} ({event['start'].get('dateTime')})" 
                                   for event in conflicts[:3]])
        
        response = self.generate_conversational_response(
            "",
            context=f"Found these conflicting events:\n{conflict_details}\nAsk user how to proceed"
        )
        self.output(response)
        
        user_choice = self.input()
        if not user_choice:
            return None
        
        if 'reschedule' in user_choice.lower():
            return self.suggest_alternative_times()
        elif 'prioritize' in user_choice.lower():
            return 'force'
        return None

    def suggest_alternative_times(self):
        available_slots = self.get_available_slots()
        if not available_slots:
            self.output("No available slots found. Please try another time.")
            return None
        
        requested_time = self.parse_time(f"{self.meeting_details['day']} {self.meeting_details['time']}")
        if not requested_time:
            return None
        
        available_slots.sort(key=lambda x: abs((x['start'] - requested_time).total_seconds()))
        
        suggestions = []
        for slot in available_slots[:3]:
            suggestions.append(slot['start'].strftime("%A at %I:%M %p"))
        
        response = self.generate_conversational_response(
            "",
            context=f"Suggest these alternative times: {', '.join(suggestions)}"
        )
        self.output(response)
        
        user_choice = self.input()
        if user_choice and any(str(i+1) in user_choice for i in range(len(suggestions))):
            selected_idx = next((i for i in range(len(suggestions)) if str(i+1) in user_choice), 0)
            selected_slot = available_slots[selected_idx]
            self.meeting_details['time'] = selected_slot['start'].strftime("%H:%M")
            self.meeting_details['day'] = selected_slot['start'].strftime("%A").lower()
            return True
        return False

    def contextual_followups(self):
        if len(self.meeting_details['attendees']) > 0:
            response = self.generate_conversational_response(
                "",
                context="Ask if user wants to invite anyone else"
            )
            self.output(response)
            
            additional_attendees = self.input()
            if additional_attendees and 'yes' in additional_attendees.lower():
                self.output("Who else should I invite? (Enter email addresses separated by commas)")
                new_attendees = self.input()
                if new_attendees:
                    self.meeting_details['attendees'].extend(
                        [email.strip() for email in new_attendees.split(',')]
                    )
        
        response = self.generate_conversational_response(
            "",
            context="Ask if user wants to set custom reminders"
        )
        self.output(response)
        
        reminder_choice = self.input()
        if reminder_choice and 'yes' in reminder_choice.lower():
            self.output("When should I remind you? (e.g., '10 minutes before' or '1 hour before')")
            reminder_time = self.input()
            if reminder_time:
                self.meeting_details['reminders'].append(reminder_time)
        
        for attendee in self.meeting_details['attendees']:
            if attendee not in self.user_preferences['usual_meeting_times']:
                self.user_preferences['usual_meeting_times'][attendee] = {}
            
            meeting_time_key = f"{self.meeting_details['day']} {self.meeting_details['time']}"
            self.user_preferences['usual_meeting_times'][attendee][meeting_time_key] = \
                self.user_preferences['usual_meeting_times'][attendee].get(meeting_time_key, 0) + 1
        
        self.save_user_preferences()

    def parse_meeting_details(self, user_input):
        prompt = """Extract meeting details from this input and return ONLY valid JSON:
{
  "duration_minutes": number,
  "day": "today|tomorrow|Monday|etc",
  "time": "HH:MM",
  "title": "string",
  "attendees": ["email1", "email2"]
}

Input: """ + user_input
        
        try:
            response = self.llm.generate_content(
                prompt,
                generation_config={"response_mime_type": "application/json"}
            )
            if response.text:
                return json.loads(response.text)
            return None
        except Exception as e:
            print(f"LLM parsing error: {e}")
            return None

    def initialize_speech_engine(self):
        try:
            self.engine = pyttsx3.init('sapi5')
            voices = self.engine.getProperty('voices')
            self.engine.setProperty('voice', voices[1].id) # 0 - male and 1 - female
            self.engine.setProperty('rate', 150)
            self.recognizer = sr.Recognizer()
        except Exception as e:
            print(f"Failed to initialize speech engine: {str(e)}")
            self.output_mode = 'text'

    def authenticate_calendar(self):
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                except Exception as e:
                    print(f"Refresh failed: {str(e)}")
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    self.creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())
        
        try:
            self.calendar_service = build('calendar', 'v3', credentials=self.creds)
        except Exception as e:
            print(f"Failed to build calendar service: {str(e)}")
            raise

    def output(self, message):
        print(f"Assistant: {message}")
        if self.output_mode == 'speech' and self.engine:
            self.engine.say(message)
            self.engine.runAndWait()
    
    def input(self, prompt=None):
        if prompt:
            self.output(prompt)
        
        if self.input_mode == 'speech':
            return self.listen()
        else:
            try:
                return input("You: ").strip()
            except EOFError:
                return None
    
    def listen(self):
        if not self.recognizer:
            self.initialize_speech_engine()
            
        with sr.Microphone() as source:
            print("Listening...")
            self.recognizer.pause_threshold = 1
            self.recognizer.adjust_for_ambient_noise(source)
            try:
                audio = self.recognizer.listen(source, timeout=5)
            except sr.WaitTimeoutError:
                self.output("I didn't hear anything. Please try again.")
                return None
        
        try:
            print("Recognizing...")
            query = self.recognizer.recognize_google(audio, language='en-in')
            print(f"User: {query}")
            return query
        except Exception as e:
            self.output("Sorry, I didn't catch that. Could you please repeat?")
            return None

    def select_interaction_mode(self):
        print("\nWelcome to Smart Scheduler Assistant!")
        print("How would you like to interact with me?")
        print("1. Text input/output")
        print("2. Voice input/output")
        print("3. Voice input with text output")
        print("4. Text input with voice output")
        
        while True:
            choice = input("Enter your choice (1-4): ").strip()
            if choice == '1':
                self.input_mode = 'text'
                self.output_mode = 'text'
                break
            elif choice == '2':
                self.input_mode = 'speech'
                self.output_mode = 'speech'
                self.initialize_speech_engine()
                break
            elif choice == '3':
                self.input_mode = 'speech'
                self.output_mode = 'text'
                self.initialize_speech_engine()
                break
            elif choice == '4':
                self.input_mode = 'text'
                self.output_mode = 'speech'
                self.initialize_speech_engine()
                break
            else:
                print("Invalid choice. Please enter 1, 2, 3, or 4.")
        
        self.output("\nGreat! I'm ready to help you schedule your meetings.")

    def parse_time(self, time_str):
        try:
            parsed_time = dateparser.parse(
                time_str,
                settings={
                    'PREFER_DATES_FROM': 'future',
                    'RELATIVE_BASE': datetime.datetime.now(pytz.timezone(TIMEZONE)),
                    'TIMEZONE': TIMEZONE
                }
            )
            if parsed_time:
                return parsed_time.astimezone(pytz.timezone(TIMEZONE))
            return None
        except Exception as e:
            print(f"Time parsing error: {str(e)}")
            return None

    def get_available_slots(self):
        day = self.parse_time(self.meeting_details.get('day', 'today')) or datetime.datetime.now(pytz.timezone(TIMEZONE))
        start_of_day = day.replace(hour=9, minute=0, second=0, microsecond=0)
        end_of_day = day.replace(hour=17, minute=0, second=0, microsecond=0)
        
        try:
            events_result = self.calendar_service.events().list(
                calendarId='primary',
                timeMin=start_of_day.isoformat(),
                timeMax=end_of_day.isoformat(),
                singleEvents=True,
                orderBy='startTime'
            ).execute()
            
            events = events_result.get('items', [])
            
            all_slots = []
            current_time = start_of_day
            while current_time < end_of_day:
                all_slots.append({
                    'start': current_time,
                    'end': current_time + timedelta(minutes=30)
                })
                current_time += timedelta(minutes=30)
            
            available_slots = []
            for slot in all_slots:
                conflict = False
                for event in events:
                    event_start_str = event['start'].get('dateTime', event['start'].get('date'))
                    event_end_str = event['end'].get('dateTime', event['end'].get('date'))
                    
                    event_start = parser.parse(event_start_str).astimezone(pytz.timezone(TIMEZONE))
                    event_end = parser.parse(event_end_str).astimezone(pytz.timezone(TIMEZONE))
                    
                    if not (slot['end'] <= event_start or slot['start'] >= event_end):
                        conflict = True
                        break
                
                if not conflict:
                    available_slots.append(slot)
            
            return available_slots
        except Exception as e:
            print(f"Error getting available slots: {str(e)}")
            return None

    def schedule_meeting(self):
        if not all([self.meeting_details['duration'], self.meeting_details['time'], self.meeting_details['day']]):
            return None
        
        requested_time = self.parse_time(f"{self.meeting_details['day']} {self.meeting_details['time']}")
        if not requested_time:
            return None
        
        available_slots = self.get_available_slots()
        if not available_slots:
            return None
        
        best_slot = None
        min_diff = float('inf')
        for slot in available_slots:
            time_diff = abs((slot['start'] - requested_time).total_seconds())
            if time_diff < min_diff:
                min_diff = time_diff
                best_slot = slot
        
        if best_slot:
            start_time = best_slot['start']
            end_time = start_time + timedelta(minutes=self.meeting_details['duration'])
            
            event = {
                'summary': self.meeting_details['title'] or 'Meeting',
                'description': 'Scheduled by Smart Scheduler',
                'start': {
                    'dateTime': start_time.isoformat(),
                    'timeZone': TIMEZONE,
                },
                'end': {
                    'dateTime': end_time.isoformat(),
                    'timeZone': TIMEZONE,
                },
                'attendees': [{'email': a} for a in self.meeting_details['attendees']],
                'reminders': {
                    'useDefault': True,
                },
                'conferenceData': {
                    'createRequest': {
                        'requestId': f"meet_{start_time.timestamp()}",
                        'conferenceSolutionKey': {
                            'type': 'hangoutsMeet'
                        }
                    }
                }
            }
            
            try:
                event = self.calendar_service.events().insert(
                    calendarId='primary',
                    body=event,
                    conferenceDataVersion=1
                ).execute()
                
                return event
            except Exception as e:
                print(f"Error scheduling meeting: {str(e)}")
                return None
        return None

    def process_conversation(self, user_input):
        if not user_input:
            return None
        
        self._current_conversation_context.append({
            'input': user_input,
            'response': None,
            'timestamp': datetime.datetime.now().isoformat()
        })
        
        if any(word in user_input.lower() for word in ['switch to', 'in', 'language']):
            lang_match = re.search(r'(english|spanish|french|german|hindi|chinese)', user_input.lower())
            if lang_match:
                lang_code = {'english': 'en', 'spanish': 'es', 'french': 'fr', 
                           'german': 'de', 'hindi': 'hi', 'chinese': 'zh-cn'}.get(lang_match.group(1), 'en')
                self.meeting_details['language'] = lang_code
                self.output(f"Switched to {lang_match.group(1)}")
                return
        
        parsed_details = self.parse_meeting_details(user_input)
        if parsed_details:
            if parsed_details.get('duration_minutes'):
                self.meeting_details['duration'] = parsed_details['duration_minutes']
            if parsed_details.get('day'):
                self.meeting_details['day'] = parsed_details['day']
            if parsed_details.get('time'):
                self.meeting_details['time'] = parsed_details['time']
            if parsed_details.get('title'):
                self.meeting_details['title'] = parsed_details['title']
            if parsed_details.get('attendees'):
                self.meeting_details['attendees'] = parsed_details['attendees']
        
        required_fields = ['duration', 'time', 'day']
        missing_fields = [field for field in required_fields if not self.meeting_details.get(field)]
        
        if not missing_fields:
            usual_time = self.suggest_usual_times(self.meeting_details['attendees'])
            if usual_time:
                response = self.generate_conversational_response(
                    user_input,
                    context=f"Suggest usual meeting time: {usual_time}"
                )
                self.output(response)
                
                use_usual_time = self.input()
                if use_usual_time and 'yes' in use_usual_time.lower():
                    time_day = self.extract_time_day(usual_time)
                    if time_day:
                        self.meeting_details.update(time_day)
            
            conflicts = self.check_for_conflicts()
            if conflicts:
                resolution = self.handle_conflict_resolution(conflicts)
                if resolution == 'force':
                    pass
                elif resolution is None:
                    return
                elif resolution is True:
                    pass
            
            confirmation = self.generate_conversational_response(
                user_input,
                context="All meeting details are complete. Ask user to confirm before scheduling."
            )
            self.output(confirmation)
            
            response = self.input()
            if response and ('yes' in response.lower() or 'confirm' in response.lower()):
                event = self.schedule_meeting()
                if event:
                    meet_link = event.get('hangoutLink', '')
                    if meet_link:
                        self.output(f"Meeting scheduled successfully! Here's your Google Meet link: {meet_link}")
                    else:
                        self.output("Meeting scheduled successfully!")
                    
                    self.contextual_followups()
                    
                    self.meeting_details = {
                        'duration': None,
                        'time': None,
                        'day': None,
                        'title': None,
                        'attendees': [],
                        'reminders': [],
                        'language': self.meeting_details['language']
                    }
                else:
                    self.output("I couldn't schedule the meeting. Please try again.")
            elif response and ('no' in response.lower() or 'cancel' in response.lower()):
                self.output("Okay, I won't schedule the meeting.")
                self.meeting_details = {
                    'duration': None,
                    'time': None,
                    'day': None,
                    'title': None,
                    'attendees': [],
                    'reminders': [],
                    'language': self.meeting_details['language']
                }
            else:
                return self.process_conversation(response)
        else:
            response = self.generate_conversational_response(
                user_input,
                context=f"Still need information about: {', '.join(missing_fields)}"
            )
            self.output(response)
        
        if self._current_conversation_context:
            self._current_conversation_context[-1]['response'] = response

    def extract_time_day(self, command):
        day_match = re.search(r'(today|tomorrow|monday|tuesday|wednesday|thursday|friday|saturday|sunday)', command, re.IGNORECASE)
        day = day_match.group(1) if day_match else None
        
        time_match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*(am|pm)?', command, re.IGNORECASE)
        if time_match:
            hour = int(time_match.group(1))
            minute = int(time_match.group(2)) if time_match.group(2) else 0
            meridian = time_match.group(3).lower() if time_match.group(3) else None
            
            if meridian:
                if meridian == 'pm' and hour < 12:
                    hour += 12
                elif meridian == 'am' and hour == 12:
                    hour = 0
            time_str = f"{hour:02d}:{minute:02d}"
            return {'day': day, 'time': time_str}
        return None

    def run(self):
        self.select_interaction_mode()
        
        self.output("""
Hello! I'm your Smart Scheduler assistant By Pranav. 
I can help you schedule meetings with Google Meet links. 
You can say things like:
- "Schedule a 30 minute meeting with john@example.com tomorrow at 2pm"
- "Set up a 1 hour team meeting called 'Project Update' on Friday"
- "I need to meet with Sarah for 45 minutes today"

What would you like to schedule?
""")
        
        while True:
            user_input = self.input()
            if self.process_conversation(user_input) == 'exit':
                break

# Constants
SCOPES = [
    'https://www.googleapis.com/auth/calendar.events',
    'https://www.googleapis.com/auth/calendar'
]
TIMEZONE = 'Asia/Kolkata'  # Change to your timezone

if __name__ == "__main__":
    scheduler = SmartScheduler()
    scheduler.run()