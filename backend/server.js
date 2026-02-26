const express = require("express");
const cors = require("cors");
const path = require("path");

const {
    Document,
    Packer,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    TextRun,
    WidthType
} = require("docx");

const app = express();

/* ------------------ MIDDLEWARE ------------------ */

app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

/* ------------------ UTILITIES ------------------ */

// Convert A/B/C/D/E → 1/2/3/4/5
function convertAnswer(ans) {
    const mapping = {
        A: "1", B: "2", C: "3", D: "4", E: "5",
        a: "1", b: "2", c: "3", d: "4", e: "5"
    };

    ans = ans.replace(/\(\d+\)/g, "").trim();
    return mapping[ans] || ans;
}

// Clean spacing but preserve line breaks
function cleanText(text) {
    return text
        .replace(/\r/g, "")
        .replace(/[ \t]+/g, " ")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

// Convert multiline string to proper DOCX paragraphs
function createMultilineParagraph(text) {
    return String(text)
        .split("\n")
        .map(line =>
            new Paragraph({
                children: [new TextRun(line)]
            })
        );
}

/* ------------------ MAIN ROUTE ------------------ */

app.post("/generate-doc", async (req, res) => {
    try {
        const content = req.body.text;

        if (!content || !content.trim()) {
            return res.status(400).json({ error: "No text provided" });
        }

        // ✅ CORRECT QUESTION SPLIT (preserves numbering)
        const questionBlocks =
            content.match(/\d+\.\s+[\s\S]*?(?=\n\d+\.\s+|$)/g) || [];

        const children = [];

        questionBlocks.forEach(block => {

            const lines = block.split("\n");

            let questionLines = [];
            let options = [];
            let answer = "";
            let solutionLines = [];
            let insideSolution = false;

            lines.forEach(rawLine => {

                let line = rawLine.trim();
                if (!line) return;

                // OPTION detection: A. A) (A)
                if (/^(\(?[A-Ea-e]\)|[A-Ea-e][\.\)])\s*/.test(line)) {
                    const optionText = line.replace(/^(\(?[A-Ea-e]\)|[A-Ea-e][\.\)])\s*/, "");
                    options.push(optionText);
                }

                // ANSWER detection
                else if (/^answer/i.test(line)) {
                    const ans = line.replace(/answer\s*[:\-]?\s*/i, "");
                    answer = convertAnswer(ans);
                }

                // SOLUTION detection
                else if (/^solution/i.test(line)) {
                    insideSolution = true;
                    const sol = line.replace(/solution\s*[:\-]?\s*/i, "");
                    solutionLines.push(sol);
                }

                else if (insideSolution) {
                    solutionLines.push(line);
                }

                // QUESTION text (preserve everything else exactly)
                else {
                    questionLines.push(rawLine);
                }
            });

            const question = cleanText(questionLines.join("\n"));
            const solution = cleanText(solutionLines.join("\n"));

            if (!question) return;

            // Support up to 5 options
            while (options.length < 5) {
                options.push("None");
            }

            const rowsData = [
                ["Question", question],
                ["Type", "Multiple Choice"],
                ["Option 1", options[0]],
                ["Option 2", options[1]],
                ["Option 3", options[2]],
                ["Option 4", options[3]],
                ["Option 5", options[4]],
                ["Answer", answer],
                ["Solution", solution],
                ["Positive Marks", "1"],
                ["Negative Marks", "0"]
            ];

            const tableRows = rowsData.map(row =>
                new TableRow({
                    children: row.map(cell =>
                        new TableCell({
                            width: {
                                size: 50,
                                type: WidthType.PERCENTAGE
                            },
                            children: createMultilineParagraph(cell)
                        })
                    )
                })
            );

            const table = new Table({
                rows: tableRows,
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE
                }
            });

            children.push(table);
            children.push(new Paragraph(""));
        });

        const doc = new Document({
            sections: [{ children }]
        });

        const buffer = await Packer.toBuffer(doc);

        res.setHeader(
            "Content-Disposition",
            "attachment; filename=MCQ_Output.docx"
        );

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );

        res.send(buffer);

    } catch (error) {
        console.error("Server Error:", error);
        res.status(500).json({ error: "Internal Server Error" });
    }
});

/* ------------------ SERVER ------------------ */

const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
    console.log(`✅ Server running on port ${PORT}`);
});