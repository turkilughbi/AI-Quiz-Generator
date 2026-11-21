
## Setup on Windows (VS Code)

1. Extract this ZIP and open the **my_Project** folder in VS Code.
2. In the terminal, run:

   ```powershell
   py -m pip install -r requirements.txt
   ```
3. change `.env.example` to `.env` iside `.env`paste a **new, private** OpenAI API key:

   ```dotenv
   OPENAI_API_KEY=YOUR_NEW_KEY_HERE
   OPENAI_MODEL=gpt-4o-mini
   ```
4. Start the server:

   ```powershell
   py app.py
   ```
5. Open **http://127.0.0.1:5000/** in your browser. Keep the terminal open while using the app.

The root URL serves the web interface; `/health` returns server health. No frontend build tools are required.

## New features

- **Easy / Medium / Hard** difficulty shown above number of questions. Hard emphasizes nuanced comparisons, application and multi-step reasoning.
- **Adaptive difficulty** checkbox: generates a larger pool with varied difficulty levels and chooses the next question based on recent correctness. This changes the order and level of questions during the attempt; it is not a separate live AI call after every answer.
- **Independent AI question validation**: a second OpenAI request reviews every question against the supplied passages, checks all four choices, discards ambiguous or unsupported questions and retries generation if too few pass. Not infallible.
- **Smart randomization and repeat reduction**: content is extracted from across the document, then sampled from varied sections for each attempt. Earlier questions are compared against recent history to reduce near repeats; repetition is allowed eventually for limited study content.
- **Document scope / focus**: choose a discovered section (PDF page groups, Word headings, or PowerPoint slides) or enter a specific topic. Section detection is structural and not a guaranteed semantic chapter index.
- **Whole-document processing**: all extractable text is indexed into chunks rather than truncating after 10,000 characters. To control cost, each quiz samples chunks from across the full source, not all pages in a single API request.
- **Optional exam timer**: checkbox enables **2 or 3 minutes per question**, with a countdown and automatic submission at zero.
- **Exam Mode**: hides correctness and explanations until submission. Switch off for immediate feedback during study.
- **Source reference and detailed feedback**: shows the original page/slide/section reference plus why every choice is right or wrong. These are model-reviewed explanations, not independently guaranteed facts.
- **Smart Review**: generate another quiz that prioritizes topics answered incorrectly in earlier attempts.
- **Local progress dashboard**: score history and per-topic accuracy stored in `quiz_history.sqlite3`; reset history for a document when needed.
- **Retake the same quiz** available separately from generating a new randomized quiz.

## How to use

Select a supported document, choose the difficulty, question count, a chapter/topic if desired, and optional exam settings. Click **Generate Quiz**. For adaptive mode, choose one answer for each question; the next question is drawn from an adaptive question pool. Select **Submit Quiz** when done, or let the timer submit automatically.

The **View Progress** button displays results for the most recently selected file. **Review Weak Topics** generates a new quiz with extra focus on previous mistakes.

## Supported files and constraints

- **PDF**, **DOCX**, **PPTX**, **TXT**, up to **15 MB** per file.
- The file must contain selectable, extractable text. **Scanned/image-only PDFs and JPG/PNG are not supported** until OCR is added.
- A document with insufficient factual material may not yield 20 distinct, validated questions; the app will ask you to try fewer questions rather than show unchecked questions.
- The AI validator uses additional API calls and will consume extra OpenAI credits. Very large documents can take longer.
- File content and sampled excerpts are submitted to the configured OpenAI model for generation and validation. Avoid uploading private documents without permission.
- For the hosted demo, each browser gets an anonymous signed visitor session (persistent for up to 90 days). Results for the same book are isolated by visitor ID and document ID. No login or cloud synchronization is included.
- Dynamic page references can be precise for PDFs and slides; Word section references depend on Heading styles. No guarantee of perfect citation alignment.
- The model instructions guide quality, but do not mathematically guarantee a single correct answer. Human review is still appropriate for important assessments.


