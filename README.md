# AppScripts
This document lists the functions available in the Google Apps Script files in this repository.

**Gemini.gs**

This file contains a GeminiClient class to interact with the Gemini API.

```GeminiClient``` 

- A class for making requests to the Gemini API.</br>

```constructor()``` 

- Initializes the client, retrieves the API key from user properties, or prompts the user for it if not found

```generateContent(prompt, temperature)``` 

- Sends a prompt to the Gemini model and returns the generated content.
- It includes retry logic for API calls.

```createGemini()``` 

- A factory function that creates and returns a new GeminiClient instance.
</br>

**sheet_to_slide.gs**

This file contains a function to generate Google Slides from a Google Sheet.

```generateSlides(config```) 

- Generates slides based on data from a spreadsheet.
- The config object specifies the source sheet, data ranges, target presentation, and template slides.
</br>

**find.email.gs**
This file contains a function to find emails in Gmail.

```find_emails(keyword, numThreads, max_length)``` 

- Searches for emails in Gmail with a subject containing the provided keyword.
- It returns the content of the found emails as a string, truncated if it exceeds the max_length

**sidebar.gs**
This file contains a function to update the sidebar in a Google Sheet.

```UpdateSideBar(msg)```
- Displays a sidebar in the spreadsheet with the provided message
