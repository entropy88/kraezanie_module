const docxLib = window.docx;

const Document = docxLib.Document || docxLib.docx?.Document;
const Packer = docxLib.Packer || docxLib.docx?.Packer;
const Paragraph = docxLib.Paragraph || docxLib.docx?.Paragraph;
const TextRun = docxLib.TextRun || docxLib.docx?.TextRun;
// ============================
// Helper Functions
// ============================

function getProps(entry, propName) {

  const regex = new RegExp(
    `${propName}\\s*=\\s*\\{([\\s\\S]*?)\\}\\s*,?`,
    "gi"
  );

  let match;
  const values = [];

  while ((match = regex.exec(entry)) !== null) {
    values.push(match[1].trim());
  }

  return values;
}

function getFirst(entry, propName) {
  const vals = getProps(entry, propName);
  return vals.length > 0 ? vals[0] : "";
}

function getJoined(entry, propName, sep = "; ") {
  const vals = getProps(entry, propName);
  const unique = [...new Set(vals)];
  return unique.join(sep);
}

function cleanNotes(entry) {
  const notesRaw = getProps(entry, "abstract");
  return notesRaw.map((n) =>
    n.replace(/^Съдържа и:\s*/i, "").trim()
  );
}

function normalizeAuthorName(name) {
  if (name.includes(",")) {
    const parts = name.split(",").map(p => p.trim());

    if (parts.length === 2) {
      return `${parts[1]} ${parts[0]}`;
    }
  }

  return name.trim();
}

function formatAuthors(rawAuthors) {
  if (!rawAuthors) return "";

  return rawAuthors
    .split(" and ")
    .map(a => normalizeAuthorName(a))
    .join(", ");
}

// ============================
// Formatting Functions
// ============================

function formatBook(entry) {
  const main_sig = getFirst(entry, "signature") || "";
  console.log(main_sig)
  const sortWord = getFirst(entry, "sort_word") || getFirst(entry, "author") || "";

  const rawAuthor = getFirst(entry, "author") || "";
  const author = formatAuthors(rawAuthor);

  let title = getFirst(entry, "title") || "";

  const subtitle = getFirst(entry, "subtitle") || getFirst(entry, "substitle") || "";
  const responsibility = getFirst(entry, "responsibility") || author;
  const edition = getJoined(entry, "edition");
  const place = getFirst(entry, "address") || getFirst(entry, "place") || "";
  const publisher = getFirst(entry, "publisher") || "";
  const year = getFirst(entry, "year") || "";
  const extent = getFirst(entry, "page_count") || getFirst(entry, "extent") || "";
  const dimensions = getFirst(entry, "illustrations") || getFirst(entry, "dimensions") || "";
  const series = getFirst(entry, "series") || "";
  const isbn = getFirst(entry, "isbn") || "";
  const book_info = getFirst(entry, "book_info") || "";

  const otherSources = getProps(entry, "other_sources");
  const itemTypes = [...new Set(getProps(entry, "item_type").map(v => v.toUpperCase()))];
  const notes = cleanNotes(entry);

  const line1 = itemTypes.includes("GOI") ? "" : sortWord;
  const line01 = `${main_sig}`;

  if (itemTypes.includes("CDD")) {
    title += " [CD-ROM]";
  }

  let line2 = `  ${title}`;

  if (subtitle) line2 += ` : ${subtitle}`;
  if (responsibility) line2 += ` / ${responsibility}`;
  if (edition) line2 += `. – ${edition}`;
  if (book_info) line2 += `. – ${book_info}`;

  if (place || publisher || year) {
    let pubBlock = "";

    if (place) pubBlock += place;
    if (publisher) pubBlock += (pubBlock ? " : " : "") + publisher;
    if (year) pubBlock += (pubBlock ? ", " : "") + year;

    line2 += `. – ${pubBlock}`;
  }

  if (extent || dimensions) {
    let physBlock = "";

    if (extent) physBlock += extent;
    if (dimensions) physBlock += (physBlock ? " ; " : "") + dimensions;

    line2 += `. – ${physBlock}`;
  }

  if (series) line2 += `. – (${series})`;
  if (isbn) line2 += `. – ISBN ${isbn}`;

  const itemTypeLine = itemTypes.length > 0
    ? `Item types: ${itemTypes.join(", ")}`
    : "";

  return {
    mainLine: `${line01}\n${line1}\n${line2}`,
    notes,
    itemTypeLine,
    otherSources,
  };
}

function formatArticle(entry) {
    
  const itemTypes = getProps(entry, "item_type")
    .map(v => v.toUpperCase());

  const sortWord = getFirst(entry, "sort_word") || "";

  const line1 = itemTypes.includes("GOI") ? "" : sortWord;

  const rawAuthor = getFirst(entry, "author") || "";
  const authorStr = formatAuthors(rawAuthor);

  let title = getFirst(entry, "title") || "";

  const source = getFirst(entry, "source") || "";
  const issue = getFirst(entry, "issue") || "";
  const year = getFirst(entry, "year") || "";
  const pages =
  getFirst(entry, "pages_art") ||
  getFirst(entry, "art_pages") ||
  "";
  const column = getFirst(entry, "column") || "";
  const journalCity = getFirst(entry, "journal_city") || "";

  const otherSources = getProps(entry, "other_sources");
  const notes = cleanNotes(entry);

  let line2 = `  ${title}`;

  if (authorStr) {
    line2 += ` / ${authorStr}`;

    if (column) {
      line2 += `.(${column})`;
    }
  }

  if (source) line2 += `. - В: ${source}`;
  if (journalCity) line2 += ` (${journalCity})`;
  if (issue) line2 += ` , бр. ${issue}`;
  if (year) line2 += ` , (${year})`;
  if (pages) line2 += ` , с. ${pages}`;

  const itemTypeLine = itemTypes.length > 0
    ? `Item types: ${itemTypes.join(", ")}`
    : "";

  return {
    mainLine: `${line1}\n${line2}`,
    notes,
    itemTypeLine,
    otherSources,
  };
}

function formatOther(entry) {
  const rawAuthor = getFirst(entry, "author") || "";
  const author = formatAuthors(rawAuthor);

  const title = getFirst(entry, "title") || "";
  const year = getFirst(entry, "year") || "";

  const itemTypes = [...new Set(getProps(entry, "item_type")
    .map(v => v.toUpperCase()))];

  const notes = cleanNotes(entry);

  let line1 = `${author}`;
  let line2 = `  ${title}`;

  if (year) line2 += ` (${year})`;

  const itemTypeLine = itemTypes.length > 0
    ? `Item types: ${itemTypes.join(", ")}`
    : "";

  return {
    mainLine: `${line1}\n${line2}`,
    notes,
    itemTypeLine,
    otherSources: [],
  };
}

// ============================
// DOCX Generation
// ============================

async function generateDocx(rawData, filename) {

  const entries = rawData
    .split(/(?=^@)/m)
    .map(e => e.trim())
    .filter(Boolean);

const parsed = entries.map((entry) => {

  const itemTypes = getProps(entry, "item_type")
    .map(v => v.toUpperCase());

  // Article if it has a source field
  const hasSource = getFirst(entry, "source");

  const recordType = hasSource
    ? "ARTICLE"
    : "BOOK";

  const year = parseInt(getFirst(entry, "year"), 10) || 0;

  return {
    entry,
    itemTypes,
    recordType,
    year,
  };
});

  
parsed.sort((a, b) => {

  const aIsBook = a.recordType === "BOOK";
  const bIsBook = b.recordType === "BOOK";

  if (aIsBook && !bIsBook) return -1;
  if (!aIsBook && bIsBook) return 1;

  return a.year - b.year;
});

const books = parsed.filter((p) =>
  p.recordType === "BOOK"
);

const articles = parsed.filter((p) =>
  p.recordType === "ARTICLE"
);

const others = [];


  const total = parsed.length;
  const children = [];

  children.push(
    new Paragraph({
      children: [
        new TextRun({
         text: "Общо записи: " + total + " (Книги: " + books.length + ", Статии: " + articles.length + ", Други: " + others.length + ")",
          bold: true,
          size: 24,
        }),
      ],
    }),
    new Paragraph({ text: "" })
  );

  function pushEntry(formatFn, entries) {

    entries.forEach((e) => {
       
      const {
        mainLine,
        notes,
        itemTypeLine,
        otherSources,
      } = formatFn(e.entry);

      mainLine.split("\n").forEach((line) => {

        const sourceMatch = line.match(/В: ([^,]+)/);

        if (sourceMatch) {

          const beforeSource = line.slice(0, sourceMatch.index + 3);
          const sourceText = sourceMatch[1];
          const afterSource = line.slice(
            sourceMatch.index + 3 + sourceText.length
          );

          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: beforeSource,
                  size: 24,
                }),
                new TextRun({
                  text: sourceText,
                  italics: true,
                  size: 24,
                }),
                new TextRun({
                  text: afterSource,
                  size: 24,
                }),
              ],
            })
          );

        } else {

          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: line,
                  size: 24,
                }),
              ],
            })
          );
        }
      });

      if (itemTypeLine) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: itemTypeLine,
                size: 24,
              }),
            ],
          })
        );
      }

      if (notes) {
        notes.forEach((n) => {
          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: "\t" + n,
                  italics: true,
                  size: 24,
                }),
              ],
            })
          );
        });
      }

      if (otherSources && otherSources.length > 0) {

        otherSources.forEach((src, idx) => {

          const prefix = idx === 0
            ? "\tВж. и: "
            : "\t";

          children.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: `${prefix}${src}`,
                  size: 24,
                }),
              ],
            })
          );
        });
      }

      children.push(new Paragraph({ text: "" }));
    });
  }

  if (books.length > 0) {

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "КНИГИ",
            bold: true,
            size: 24,
          }),
        ],
      })
    );

    pushEntry(formatBook, books);
  }

  if (articles.length > 0) {

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "СТАТИИ",
            bold: true,
            size: 24,
          }),
        ],
      })
    );

    pushEntry(formatArticle, articles);
  }

  if (others.length > 0) {

    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "ДРУГИ",
            bold: true,
            size: 24,
          }),
        ],
      })
    );

    pushEntry(formatOther, others);
  }

  const doc = new Document({
    sections: [{ children }],
  });

  const blob = await Packer.toBlob(doc);

  saveAs(blob, `${filename}.docx`);
}

// ============================
// UI
// ============================

const generateBtn = document.getElementById("generateBtn");
const statusDiv = document.getElementById("status");

generateBtn.addEventListener("click", async () => {

  const fileInput = document.getElementById("bibfile");

  const filename = document.getElementById("filename")
    .value
    .trim() || "sorted_shelf";

  if (!fileInput.files.length) {
    statusDiv.textContent = "Please select a BibTeX file.";
    return;
  }

  try {

    statusDiv.textContent = "Reading file...";

    const file = fileInput.files[0];

    const rawData = await file.text();

    statusDiv.textContent = "Generating DOCX...";

    await generateDocx(rawData, filename);

    statusDiv.textContent = "Done!";

  } catch (err) {

    console.error(err);

    statusDiv.textContent = "Error generating DOCX.";
  }
});
