import os
import re
import json
import pytz
import sys
import pickle
import datetime
from datetime import timedelta
import dateparser
from dateutil import parser
import speech_recognition as sr
import pyttsx3
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold

class SmartScheduler:
    def __init__(self):
        self.input_mode = None
        self.output_mode = None
        self.creds = None
        self.calendar_service = None
        self.conversation_history = []
        self.meeting_context = {
            'purpose': None,
            'duration': None,
            'time_preference': None,
            'day_preference': None,
            'attendees': [],
            'constraints': []
        }
        self.user_preferences = self.load_user_preferences()
        self.engine = None
        self.recognizer = None
        self.llm = self.initialize_gemini()
        self.timezone = 'Asia/Kolkata'
        
        try:
            self.authenticate_calendar()
        except Exception as e:
            print(f"Authentication failed: {str(e)}")
            sys.exit(1)

    def initialize_gemini(self):
        genai.configure(api_key='AIzaSyCSbCZ1wem49jxmQAJZYVTzk6rHBDlIqgs')
        return genai.GenerativeModel('gemini-1.5-flash-latest')

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
            'default_reminders': ['email', 10],
            'favorite_times': ['10:00', '14:00', '15:30']
        }

    def authenticate_calendar(self):
        SCOPES = ['https://www.googleapis.com/auth/calendar']
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())
        
        self.calendar_service = build('calendar', 'v3', credentials=self.creds)

    def initialize_speech_engine(self):
        try:
            self.engine = pyttsx3.init('sapi5')
            voices = self.engine.getProperty('voices')
            self.engine.setProperty('voice', voices[1].id)
            self.engine.setProperty('rate', 150)
            self.recognizer = sr.Recognizer()
        except Exception as e:
            print(f"Failed to initialize speech engine: {str(e)}")
            self.output_mode = 'text'

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

    def parse_time(self, time_str, reference_date=None):
        """Enhanced time parsing with Asia/Kolkata timezone"""
        try:
            settings = {
                'PREFER_DATES_FROM': 'future',
                'RELATIVE_BASE': datetime.datetime.now(pytz.timezone(self.timezone)),
                'TIMEZONE': self.timezone,
                'RETURN_AS_TIMEZONE_AWARE': True
            }
            
            if reference_date:
                settings['RELATIVE_BASE'] = reference_date
            
            parsed_time = dateparser.parse(time_str, settings=settings)
            
            if parsed_time:
                # Handle relative days (e.g., "next Tuesday")
                if not any(word in time_str.lower() for word in ['today', 'tomorrow', 'yesterday']):
                    if parsed_time.date() < datetime.datetime.now(pytz.timezone(self.timezone)).date():
                        parsed_time += timedelta(days=7)
                
                return parsed_time.astimezone(pytz.timezone(self.timezone))
            return None
        except Exception as e:
            print(f"Time parsing error: {str(e)}")
            return None

    def resolve_relative_time(self, time_phrase):
        """Handle complex relative time references using LLM"""
        prompt = f"""Interpret this time reference and return only the specific date/time in ISO format:
Current date: {datetime.datetime.now(pytz.timezone(self.timezone)).strftime('%Y-%m-%d')}
Time zone: {self.timezone}

Time reference: "{time_phrase}"

Consider:
1. Relative dates ("next Tuesday")
2. Event-based references ("after my 2pm meeting")
3. Calendar logic ("last weekday of this month")
4. Time buffers ("1 hour before my flight")

Return ONLY the resolved date/time in ISO format or "unknown" if unclear."""
        
        try:
            response = self.llm.generate_content(prompt)
            if response.text.lower() != "unknown":
                return parser.parse(response.text).astimezone(pytz.timezone(self.timezone))
        except Exception as e:
            print(f"LLM time resolution error: {e}")
        return None

    def get_available_slots(self, day, duration_minutes=30, time_range=None):
        """Get available slots for a specific day with Asia/Kolkata working hours"""
        try:
            if isinstance(day, str):
                day = self.resolve_relative_time(day) or self.parse_time(day)
            
            if not day:
                return None
                
            # Standard working hours in India (9AM to 6PM)
            start_of_day = day.replace(hour=9, minute=0, second=0, microsecond=0)
            end_of_day = day.replace(hour=18, minute=0, second=0, microsecond=0)
            
            # Adjust for specific time ranges
            if time_range:
                start_hour, end_hour = time_range
                start_of_day = start_of_day.replace(hour=start_hour)
                end_of_day = end_of_day.replace(hour=end_hour)
            
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
                    'end': current_time + timedelta(minutes=duration_minutes)
                })
                current_time += timedelta(minutes=15)  # Check every 15 minutes
            
            available_slots = []
            for slot in all_slots:
                conflict = False
                for event in events:
                    event_start = parser.parse(event['start'].get('dateTime', event['start'].get('date'))).astimezone(pytz.timezone(self.timezone))
                    event_end = parser.parse(event['end'].get('dateTime', event['end'].get('date'))).astimezone(pytz.timezone(self.timezone))
                    
                    if not (slot['end'] <= event_start or slot['start'] >= event_end):
                        conflict = True
                        break
                
                if not conflict:
                    available_slots.append(slot)
            
            return available_slots
        except Exception as e:
            print(f"Error getting available slots: {str(e)}")
            return None

    def create_calendar_event(self, title, start_time, end_time, attendees=None):
        """Actually create the calendar event with Google Meet"""
        if not attendees:
            attendees = []
        
        event = {
            'summary': title,
            'start': {
                'dateTime': start_time.isoformat(),
                'timeZone': self.timezone,
            },
            'end': {
                'dateTime': end_time.isoformat(),
                'timeZone': self.timezone,
            },
            'attendees': [{'email': email} for email in attendees],
            'conferenceData': {
                'createRequest': {
                    'requestId': f"{title.replace(' ', '')}_{start_time.timestamp()}",
                    'conferenceSolutionKey': {'type': 'hangoutsMeet'}
                }
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'email', 'minutes': 24 * 60},
                    {'method': 'popup', 'minutes': 30},
                ],
            },
        }
        
        try:
            event = self.calendar_service.events().insert(
                calendarId='primary',
                body=event,
                conferenceDataVersion=1
            ).execute()
            return event
        except Exception as e:
            print(f"Error creating event: {e}")
            return None

    def extract_meeting_details(self, user_input):
        """Use LLM to extract structured meeting details"""
        prompt = f"""Extract meeting details from this request in JSON format:
{user_input}

Include:
- title (string)
- duration (minutes)
- day (e.g., "Tuesday")
- date (specific date if mentioned)
- time_range (e.g., "9-12")
- attendees (list of emails)
- purpose (string)

Return ONLY valid JSON with these fields."""
        
        try:
            response = self.llm.generate_content(prompt)
            return json.loads(response.text)
        except Exception as e:
            print(f"Error extracting meeting details: {e}")
            return {}

    def handle_scheduling(self, user_input):
        """Handle the complete scheduling workflow"""
        # Extract meeting details using LLM
        meeting_details = self.extract_meeting_details(user_input)
        
        # Calculate next Tuesday if requested
        if 'next tuesday' in user_input.lower():
            today = datetime.datetime.now(pytz.timezone(self.timezone))
            days_ahead = (1 - today.weekday()) % 7  # Tuesday is weekday 1
            if days_ahead <= 0:  # If today is Tuesday, get next week's Tuesday
                days_ahead += 7
            next_tuesday = today + datetime.timedelta(days_ahead)
            meeting_details['day'] = next_tuesday.strftime('%A')
            meeting_details['date'] = next_tuesday.date().isoformat()
        
        # Set default duration if not specified
        if not meeting_details.get('duration'):
            meeting_details['duration'] = 30  # Default to 30 minutes
        
        # Find available time slot
        if meeting_details.get('day'):
            slots = self.get_available_slots(
                meeting_details['day'],
                meeting_details['duration'],
                self.parse_time_range(meeting_details.get('time_range'))
            )
            
            if slots:
                chosen_slot = slots[0]
                # Actually create the event
                event = self.create_calendar_event(
                    meeting_details.get('title', 'Meeting'),
                    chosen_slot['start'],
                    chosen_slot['end'],
                    meeting_details.get('attendees', [])
                )
                
                if event:
                    meeting_link = event.get('hangoutLink', 'Google Meet link generated')
                    return (f"I've scheduled '{meeting_details.get('title', 'Meeting')}' "
                           f"on {chosen_slot['start'].strftime('%A, %B %d at %I:%M %p')} "
                           f"for {meeting_details['duration']} minutes. "
                           f"Here's your Google Meet link: {meeting_link}")
        
        return "I couldn't schedule the meeting. Please try again with more specific details."

    def parse_time_range(self, time_range_str):
        """Convert time range string to (start_hour, end_hour)"""
        if not time_range_str:
            return None
        
        try:
            start, end = map(int, time_range_str.split('-'))
            return (start, end)
        except:
            return None

    def generate_response(self, user_input):
        """Enhanced LLM response generation with smart scheduling logic"""
        # Handle date/time queries directly
        if 'today' in user_input.lower():
            today = datetime.datetime.now(pytz.timezone(self.timezone))
            return f"Today is {today.strftime('%A, %B %d, %Y')} and the current time is {today.strftime('%I:%M %p')} IST."
        
        if 'time now' in user_input.lower():
            now = datetime.datetime.now(pytz.timezone(self.timezone))
            return f"The current time is {now.strftime('%I:%M %p')} IST."
        
        # Handle scheduling requests
        if any(word in user_input.lower() for word in ['schedule', 'meeting', 'appointment']):
            return self.handle_scheduling(user_input)
        
        # Default LLM response
        prompt = f"""You are an advanced scheduling assistant in {self.timezone} timezone. Current context:
{json.dumps(self.meeting_context, indent=2)}

Conversation history:
{self.conversation_history[-5:] if self.conversation_history else 'None'}

User input: "{user_input}"

Generate a concise, helpful response that:
1. Answers time/date queries accurately
2. Handles scheduling requests efficiently
3. Provides Google Meet links when appropriate
4. Uses natural, friendly language"""
        
        try:
            response = self.llm.generate_content(prompt)
            return response.text
        except Exception as e:
            print(f"LLM error: {e}")
            return "I encountered an error. Let's try again."

    def process_conversation(self):
        self.output("Hello! I'm your Smart Scheduling Assistant. How can I help you schedule today?")
        
        while True:
            user_input = self.input()
            if not user_input:
                continue
            
            if user_input.lower() in ['exit', 'quit', 'bye']:
                self.output("Goodbye! Have a great day.")
                break
            
            self.conversation_history.append(f"User: {user_input}")
            response = self.generate_response(user_input)
            self.output(response)
            self.conversation_history.append(f"Assistant: {response}")

    def run(self):
        self.select_interaction_mode()
        self.process_conversation()

if __name__ == "__main__":
    scheduler = SmartScheduler()
    scheduler.run()