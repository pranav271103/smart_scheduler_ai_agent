# Smart Scheduler AI Assistant

A powerful AI-powered meeting scheduler that helps you manage your calendar and schedule meetings with ease. This assistant uses natural language processing, speech recognition, and Google Calendar integration to make scheduling meetings as simple as having a conversation.

## Features

- 🤖 AI-powered natural language processing for understanding meeting requests
- 🎤 Multiple interaction modes:
  - Text input/output
  - Voice input/output
  - Voice input with text output
  - Text input with voice output
- 🌍 Multilingual support (English, Spanish, French, German, Hindi, Chinese)
- 📅 Google Calendar integration
- 🎯 Smart conflict resolution
- 📊 Learning from user preferences
- 🔔 Customizable reminders
- 🎥 Automatic Google Meet link generation
- 🎯 Contextual follow-ups and suggestions

🎥 [Click here to watch the demo video](https://drive.google.com/file/d/1EkjzgMfcZW9CFCcnfxzZQcOdaCWuizLs/view?usp=sharing)

## Prerequisites

- Python 3.7 or higher
- Google Cloud Platform account with Calendar API enabled
- Google OAuth 2.0 credentials

## Installation

1. Clone the repository:
```bash
git clone https://github.com/pranav271103/smart_scheduler_ai_agent.git
cd smart_scheduler_ai_agent
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Google Calendar API Credentials:

Go to the [Google Cloud Console](https://console.cloud.google.com/).
**Create a new project** (or select an existing one).
**Enable the Google Calendar API**:
   - Navigate to **APIs & Services > Library**.
   - Search for **"Google Calendar API"**.
   - Click **Enable**.
**Create OAuth credentials**:
   - Go to **APIs & Services > Credentials**.
   - Click **Create Credentials > OAuth client ID**.
   - Choose **Desktop app** as the application type.
   - Name your client (e.g., `SmartScheduler Client`).
   - Download the JSON file.
**Rename** the file to `credentials.json`.
**Place** `credentials.json` in the **project root directory**.

---

### Gemini API Key

1. Visit [Google AI Studio](https://makersuite.google.com/app/apikey).
2. Click **"Get API Key"** in the left sidebar.
3. Create a new API key if needed.
4. Copy the generated key.
5. In your `SmartAi.py` file, locate this line:

   ```python
   genai.configure(api_key='YOUR_ACTUAL_API_KEY_HERE')


## Configuration

1. Place your `credentials.json` file in the project root directory
2. Update the `TIMEZONE` constant in `SmartAi.py` to match your timezone
3. (Optional) Update the Gemini API key in the `initialize_gemini` method

## Usage

Run the scheduler:
```bash
python SmartAi.py
```

The assistant will prompt you to select your preferred interaction mode. You can then start scheduling meetings using natural language commands like:

- "Schedule a 30 minute meeting with john@example.com tomorrow at 2pm"
- "Set up a 1 hour team meeting called 'Project Update' on Friday"
- "I need to meet with Sarah for 45 minutes today"

## Features in Detail

### Natural Language Processing
The assistant uses Google's Gemini AI model to understand and process natural language input, making it easy to schedule meetings using everyday language.

### Smart Conflict Resolution
When scheduling conflicts are detected, the assistant will:
- Show existing conflicts
- Suggest alternative times
- Allow rescheduling or prioritizing the new meeting

### Learning from Preferences
The assistant learns from your scheduling patterns and preferences:
- Remembers usual meeting times with specific attendees
- Suggests preferred times based on history
- Adapts to your scheduling habits

### Multilingual Support
Switch between languages using commands like:
- "Switch to Spanish"
- "In French"
- "Change language to Hindi"

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Google Calendar API
- Google Gemini AI
- PyTTSx3 for text-to-speech
- SpeechRecognition for voice input
- Google Translate API for multilingual support