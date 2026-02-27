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

// Middlewares
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// Serve static frontend
app.use(express.static(path.join(__dirname, "public")));

// Convert A/B/C/D/E → 1/2/3/4/5
function convertAnswer(ans) {
    const mapping = {
        A: "1", B: "2", C: "3", D: "4", E: "5",
        a: "1", b: "2", c: "3", d: "4", e: "5"
    };

    ans = ans.replace(/\(\d+\)/g, "").trim();
    return mapping[ans] || ans;
}

// Generate DOC endpoint
app.post("/generate-doc", async (req, res) => {
    try {
        const content = req.body.text;

        if (!content || !content.trim()) {
            return res.status(400).json({ error: "No text provided" });
        }

        const questionBlocks = content
            .split(/\n?\d+\.\s+/)
            .filter(q => q.trim());

        const children = [];

        questionBlocks.forEach((block, index) => {

            const lines = block.split("\n");

            let question = "";
            let options = [];
            let answer = "";
            let solution = "";
            let isSolution = false;

            lines.forEach(line => {
                line = line.trim();

                if (!line) return;

                // OPTION
                if (/^[A-Ea-e][\.\)]\s*/.test(line)) {
                    const optionText = line.replace(/^[A-Ea-e][\.\)]\s*/, "");
                    options.push(optionText);
                    return;
                }

                // ANSWER
                if (/^answer/i.test(line)) {
                    const ans = line.replace(/answer\s*[:\-]?\s*/i, "");
                    answer = convertAnswer(ans);
                    return;
                }

                // START SOLUTION
                if (/^solution/i.test(line)) {
                    isSolution = true;
                    const firstLine = line.replace(/solution\s*[:\-]?\s*/i, "");
                    if (firstLine) solution += firstLine + " ";
                    return;
                }

                // IF IN SOLUTION MODE
                if (isSolution) {
                    solution += line + " ";
                }
                else {
                    question += line + " ";
                }
            });

            question = question.trim();

            if (!question) return;

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
                            children: [
                                new Paragraph({
                                    children: [new TextRun(String(cell))]
                                })
                            ]
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
            sections: [
                {
                    children: children
                }
            ]
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

// Render-compatible port
const PORT = process.env.PORT || 5000;

app.listen(PORT, () => {
    console.log(`✅ Server running on port ${PORT}`);
});