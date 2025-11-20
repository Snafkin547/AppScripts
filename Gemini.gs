```// Configure,
GEMINI_MODEL, e.g. 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro'
MAX_RETRIES: Max retries by Gemini,  e.g. 3 
INITIAL_BACKOFF_MS: Backof time in MS, e.g. 10000
```

class GeminiClient {
  constructor() {
    const userProperties = PropertiesService.getUserProperties();
    let apiKey = userProperties.getProperty('GEMINI_API_KEY');

    if (!apiKey) {
      apiKey = Browser.inputBox(
        "Gemini API Key Setup", 
        "Please enter your Gemini API Key🔑.\nIt will be saved securely for your user account.", 
        Browser.Buttons.OK_CANCEL
      );
      
      if (apiKey === "cancel" || apiKey === "") {
        SpreadsheetApp.getUi().alert("No API Key provided. Terminating the process.");
        throw new Error("API Key not provided.");
      }
      
      userProperties.setProperty('GEMINI_API_KEY', apiKey);
      UpdateSideBar("API Key saved successfully for future use.");
    }
    
    this.apiKey = apiKey;
    this.model = GEMINI_MODEL;
    this.url = `${this.model}:generateContent?key=${this.apiKey}`;
  }

  /**
   * Generates content with built-in retry logic.
   * @param {string} prompt The prompt to send to the model.
   * @param {number} temperature The generation temperature.
   * @returns {Object} The parsed JSON object from the API.
   * @throws {Error} If the API call fails after all retries.
   */
  generateContent(prompt, temperature) {
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: temperature,
        responseMimeType: "application/json",
      },
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    // --- RETRY LOGIC ---
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      try {
        const response = UrlFetchApp.fetch(this.url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
          return JSON.parse(responseBody);
        } else {
          Logger.log(`API Error ${responseCode} on attempt ${attempt + 1}: ${responseBody}`);
          throw new Error(`API Error ${responseCode}: ${responseBody}`);
        }
      } catch (e) {
        Logger.log(`Attempt ${attempt + 1} for prompt failed: ${e.message}`);
        if (attempt === MAX_RETRIES - 1) {
          // Last attempt failed, throw the error to be caught by the calling function.
          throw new Error(`AI model failed after ${MAX_RETRIES} attempts. Last error: ${e.message}`);
        }
        // Calculate sleep time with jitter and wait.
        const sleepTime = (2 ** attempt) * INITIAL_BACKOFF_MS + Math.floor(Math.random() * 1000);
        Logger.log(`Waiting ${sleepTime / 1000}s before retrying...`);
        Utilities.sleep(sleepTime);
      }
    }
  }
}

function createGemini(){
  return new GeminiClient();
}