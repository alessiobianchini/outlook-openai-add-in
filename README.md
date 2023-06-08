<h1 align="center">
Outlook OpenAI Add-in
</h1>

<p align="center">Outlook Add-in with OpenAI APIs.</p>

---

## This Outlook add-in is developed with the React framework and uses Office.js and OpenAI packages to interact with the Outlook application and OpenAI APIs.

## Demo app

Run `npm install` and `npm start` for a dev server with an example. Test it in your Outlook desktop application.

## Usage

![image](https://github.com/alessiobianchini/outlook-openai-add-in/assets/33493281/8fc578f7-b9bb-42ae-8edd-5efcaf7e6746)


- "Generate business mail" -> Generates a business email from the selected text in the original language.

- "Translate to English" -> Translates the selected text into English.

- "Correct spelling and grammar" -> Corrects the spelling and grammar of the selected text in the original language.

- "Set OpenAI token" -> Setup your personal OpenAI api token:

  ![image](https://github.com/alessiobianchini/outlook-openai-add-in/assets/33493281/74e55ebc-a4b9-4fc3-aabe-55ed61790c16)

## Deployment

- Add your urlProd to the `webpack.config.js` file.
- Run the `npm run build` command and publish the compiled output.
- Import the Add-in from the url yoursite.com/manifest.xml.
