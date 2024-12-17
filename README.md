# NoBBQ

**We are using the following research/repo as a basis for our Norwegian version of the same thing;**
- https://github.com/nyu-mll/BBQ/blob/main/README.md


**This project is initiated by:**
- Wessel Braakman
- Alejandra Palacio Perez
- Teresa Dalen Herland

**Note that this is a PoC type of project, nothing here is set in stone and all conclusions from this project are NOT based on empirical/scientifical evidence.
We start this project, hoping it will be picked up by an institution that has the capacity to take this to a professional level.**

**Steps:**
1. Download raw JSONL files from the original BBQ repository (per category)
2. Filter these files so we end up with a maximum of 50 unique contexts/questions (per category)
3. Determine whether we can either reuse, change or have to delete the contexts or questions (looking at Norwegian society)
4. Translate the contexts and questions to Norwegian
5. Create prompts (context and questions)
6. Send prompts to various LLM systems (e.g. ChatGPT, Perplexity, Gemini, ...)
7. Document responses
8. Review the responses (do they contain bias)
9. Report conclusions
