const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
  HeadingLevel, PageBreak, Header, Footer, PageNumber
} = require("docx");

// Constants
const PAGE_WIDTH = 12240; // 8.5in
const PAGE_HEIGHT = 15840; // 11in
const MARGIN = 1440; // 1in
const CONTENT_WIDTH = PAGE_WIDTH - (2 * MARGIN); // 9360 DXA

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0 };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Scavenger hunt items
const scavengerItems = [
  "Pinecone", "Fern", "Spider web", "Smooth rock",
  "Mushroom", "Bird", "Squirrel", "Wildflower",
  "Moss", "Feather", "Ant trail", "Butterfly",
  "Bark pattern", "Snail", "Acorn", "Animal tracks"
];

// Bingo items (5x5)
const bingoItems = [
  "Marshmallow fell in fire", "Saw a shooting star", "Heard an owl", "Dog barked at nothing", "Found a cool stick",
  "Told a joke", "Saw the Big Dipper", "Heard a frog", "Roasted perfect marshmallow", "Someone said 'this is nice'",
  "Flashlight tag", "Counted 10 stars", "FREE SPACE", "Heard a splash", "Wind blew something over",
  "Made someone laugh", "Saw a bat", "Dog fell asleep first", "Smelled campfire smoke", "Found heart-shaped rock",
  "Saw the moon", "Heard a bird at night", "Made a shadow puppet", "Said goodnight to the trees", "Yawned first"
];

// Word search grid (12x12)
const wordGrid = [
  "T E N T F D S T A R S J",
  "Q S S D R P N O O T R C",
  "Y E A X B H I K E N H A",
  "H A Y A H L L A K E S M",
  "F G Y C P S M Y N T I P",
  "E L O P Y B H R A R F F",
  "R E E R T V B R H A I I",
  "R Q H B F T S E F I S R",
  "Y G C I C Y K F Y L H E",
  "R C K U C A N O E A N D",
  "E J C Q Y O W L V E F E",
  "S J I S L A N D Q F G B"
];
const wordList = ["TENT", "CAMPFIRE", "LAKE", "HIKE", "TRAIL", "SMORES", "STARS", "FISH", "CANOE", "TREES", "FROG", "OWL", "DEER", "EAGLE", "ISLAND", "FERRY"];

// Decoder key
const decoderKey = "A=Pine B=Bear C=Camp D=Deer E=Eagle F=Fire G=Fern H=Chipmunk I=Fish J=Log K=Mushroom L=Leaf M=Moon N=Snail O=Owl P=Rock Q=Flower R=Rabbit S=Star T=Canoe U=Butterfly V=Wave W=Web X=Feather Y=Paw Z=Sun";

// Helper: page break paragraph
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Helper: title for each activity page
function activityTitle(title, subtitle) {
  const children = [
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: title, bold: true, size: 36, font: "Georgia" })]
    })
  ];
  if (subtitle) {
    children.push(new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: subtitle, size: 22, italics: true, color: "555555" })]
    }));
  }
  return children;
}

// Helper: empty lines for writing
function writingLines(count) {
  const lines = [];
  for (let i = 0; i < count; i++) {
    lines.push(new Paragraph({
      spacing: { before: 60 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB", space: 1 } },
      children: [new TextRun({ text: " ", size: 22 })]
    }));
  }
  return lines;
}

// Helper: drawing box (bordered empty space)
function drawingBox(label, heightLines) {
  const rows = [];
  const cellContent = [
    new Paragraph({
      spacing: { after: 40 },
      children: [new TextRun({ text: label, size: 20, italics: true, color: "666666" })]
    })
  ];
  for (let i = 0; i < heightLines; i++) {
    cellContent.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
  }
  rows.push(new TableRow({
    children: [new TableCell({
      borders: { top: { style: BorderStyle.DASHED, size: 2, color: "999999" }, bottom: { style: BorderStyle.DASHED, size: 2, color: "999999" }, left: { style: BorderStyle.DASHED, size: 2, color: "999999" }, right: { style: BorderStyle.DASHED, size: 2, color: "999999" } },
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: cellContent
    })]
  }));
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows
  });
}

// ===== PAGE 1: SCAVENGER HUNT =====
function scavengerHuntPage() {
  const cellWidth = Math.floor(CONTENT_WIDTH / 4);
  const rows = [];
  for (let r = 0; r < 4; r++) {
    const cells = [];
    for (let c = 0; c < 4; c++) {
      const item = scavengerItems[r * 4 + c];
      cells.push(new TableCell({
        borders,
        width: { size: cellWidth, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [
          new Paragraph({
            spacing: { after: 60 },
            children: [new TextRun({ text: "\u2610 ", size: 28 }), new TextRun({ text: item, size: 22, bold: true })]
          })
        ]
      }));
    }
    rows.push(new TableRow({ children: cells }));
  }

  return [
    ...activityTitle("Nature Scavenger Hunt", "Check off every wild thing you spot around camp."),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: "Explorer mission: Find all 16 nature treasures while hiking, playing by camp, or walking with the dogs.", size: 22 })]
    }),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [cellWidth, cellWidth, cellWidth, cellWidth],
      rows
    }),
    new Paragraph({ spacing: { before: 300 }, children: [] }),
    drawingBox("Bonus: Draw your coolest find!", 6)
  ];
}

// ===== PAGE 2: CAMPFIRE BINGO =====
function bingoPage() {
  const cellWidth = Math.floor(CONTENT_WIDTH / 5);
  const rows = [];
  for (let r = 0; r < 5; r++) {
    const cells = [];
    for (let c = 0; c < 5; c++) {
      const item = bingoItems[r * 5 + c];
      const isCenter = item === "FREE SPACE";
      cells.push(new TableCell({
        borders,
        width: { size: cellWidth, type: WidthType.DXA },
        verticalAlign: "center",
        shading: isCenter ? { fill: "FFF3CD", type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: item, size: 18, bold: isCenter })]
          })
        ]
      }));
    }
    rows.push(new TableRow({ height: { value: 1200, rule: "atLeast" }, children: cells }));
  }

  return [
    ...activityTitle("Campfire Bingo", "Mark a square whenever it happens. Can you get five in a row?"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [cellWidth, cellWidth, cellWidth, cellWidth, cellWidth],
      rows
    })
  ];
}

// ===== PAGE 3: WORD SEARCH =====
function wordSearchPage() {
  const cellWidth = Math.floor((CONTENT_WIDTH * 0.7) / 12);
  const gridRows = wordGrid.map(row => {
    const letters = row.split(" ");
    return new TableRow({
      children: letters.map(letter => new TableCell({
        borders,
        width: { size: cellWidth, type: WidthType.DXA },
        verticalAlign: "center",
        margins: { top: 40, bottom: 40, left: 0, right: 0 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: letter, size: 24, bold: true, font: "Courier New" })]
        })]
      }))
    });
  });

  const wordListParas = [
    new Paragraph({ spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "Find these words:", size: 22, bold: true })] })
  ];
  for (let i = 0; i < wordList.length; i += 4) {
    const chunk = wordList.slice(i, i + 4).join("   ");
    wordListParas.push(new Paragraph({ children: [new TextRun({ text: chunk, size: 20 })] }));
  }

  return [
    ...activityTitle("Camping Word Search", "Circle the camping words. They go across, down, and diagonal."),
    new Table({
      width: { size: cellWidth * 12, type: WidthType.DXA },
      columnWidths: Array(12).fill(cellWidth),
      rows: gridRows
    }),
    ...wordListParas
  ];
}

// ===== PAGE 4-6: NATURE JOURNALS =====
function journalPage(dayNum, title, subtitle) {
  return [
    ...activityTitle(`Nature Journal - Day ${dayNum}: ${title}`, subtitle),
    drawingBox(`Draw what Day ${dayNum} looked like`, 8),
    new Paragraph({ spacing: { before: 200 }, children: [] }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Today I saw...", size: 22, bold: true, color: "2C4A1E" })] }),
    ...writingLines(2),
    new Paragraph({ spacing: { before: 120, after: 40 }, children: [new TextRun({ text: "My favorite part was...", size: 22, bold: true, color: "2C4A1E" })] }),
    ...writingLines(2),
    new Paragraph({ spacing: { before: 120, after: 40 }, children: [new TextRun({ text: "Tomorrow I want to...", size: 22, bold: true, color: "2C4A1E" })] }),
    ...writingLines(2)
  ];
}

// ===== PAGE 7: SECRET CODE DECODER =====
function decoderPage() {
  const keyEntries = [
    "A = Pine Tree", "N = Snail", "B = Bear", "O = Owl",
    "C = Campsite", "P = Rock", "D = Deer", "Q = Flower",
    "E = Eagle", "R = Rabbit", "F = Fire", "S = Star",
    "G = Fern", "T = Canoe", "H = Chipmunk", "U = Butterfly",
    "I = Fish", "V = Wave", "J = Log", "W = Web",
    "K = Mushroom", "X = Feather", "L = Leaf", "Y = Paw Print",
    "M = Moon", "Z = Sun"
  ];

  const cellWidth = Math.floor(CONTENT_WIDTH / 4);
  const keyRows = [];
  for (let r = 0; r < 7; r++) {
    const cells = [];
    for (let c = 0; c < 4; c++) {
      const idx = r * 4 + c;
      const text = idx < keyEntries.length ? keyEntries[idx] : "";
      cells.push(new TableCell({
        borders: { top: noBorder, bottom: { style: BorderStyle.DOTTED, size: 1, color: "CCCCCC" }, left: noBorder, right: noBorder },
        width: { size: cellWidth, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: true })] })]
      }));
    }
    keyRows.push(new TableRow({ children: cells }));
  }

  const messages = [
    { label: "Message 1", hint: "(I LOVE CAMPING)", code: "Fish / Leaf Owl Wave Eagle / Campsite Pine Moon Rock Fish Snail Fern" },
    { label: "Message 2", hint: "(ORCAS ISLAND)", code: "Owl Rabbit Campsite Pine Star / Fish Star Leaf Pine Snail Deer" },
    { label: "Message 3", hint: "(BEST TRIP EVER)", code: "Bear Eagle Star Canoe / Canoe Rabbit Fish Rock / Eagle Wave Eagle Rabbit" }
  ];

  const messageParas = [];
  messages.forEach(m => {
    messageParas.push(new Paragraph({
      spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: m.label, size: 22, bold: true })]
    }));
    messageParas.push(new Paragraph({
      spacing: { after: 40 },
      children: [new TextRun({ text: m.code, size: 20, italics: true })]
    }));
    messageParas.push(new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB", space: 1 } },
      spacing: { after: 40 },
      children: [new TextRun({ text: "Answer: ", size: 20 })]
    }));
  });

  return [
    ...activityTitle("Secret Code Decoder", "Use the key below to decode the secret messages!"),
    new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: "CODE KEY", size: 20, bold: true, color: "2C4A1E" })] }),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [cellWidth, cellWidth, cellWidth, cellWidth],
      rows: keyRows
    }),
    new Paragraph({ spacing: { before: 200, after: 60 }, children: [new TextRun({ text: "DECODE THESE MESSAGES", size: 20, bold: true, color: "2C4A1E" })] }),
    new Paragraph({ children: [new TextRun({ text: "Use / to mark spaces between words.", size: 18, italics: true, color: "666666" })] }),
    ...messageParas
  ];
}

// ===== PAGE 8: DOT-TO-DOT (text instructions version) =====
function dotToDotPage() {
  return [
    ...activityTitle("Dot-to-Dot: Camp Scene", "Connect the dots from 1 to 25 to reveal a tent with trees!"),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Start at dot 1 and draw a line to each number in order. When you finish, color in your picture!", size: 22 })] }),
    drawingBox("Connect the dots here (use the numbered guide below)", 12),
    new Paragraph({ spacing: { before: 200, after: 100 }, children: [new TextRun({ text: "Dot Guide (approximate positions on the page above):", size: 18, italics: true })] }),
    new Paragraph({ children: [new TextRun({ text: "Dots 1-11 form a tall pine tree on the left side", size: 18 })] }),
    new Paragraph({ children: [new TextRun({ text: "Dots 12-17 form a tent in the center", size: 18 })] }),
    new Paragraph({ children: [new TextRun({ text: "Dots 18-25 form a second tree on the right side", size: 18 })] }),
  ];
}

// ===== PAGE 9: CAMPING JOKES =====
function jokesPage() {
  const jokes = [
    ["What do you call a bear with no teeth?", "A gummy bear!"],
    ["Why did the tree pack a suitcase?", "It was ready to leaf town."],
    ["What kind of music do campfires like?", "Hot tracks."],
    ["Why do frogs like camping?", "Because they get to eat French flies."],
    ["What do you call an owl who tells jokes at camp?", "A real hoot."],
    ["Why was the tent so good at secrets?", "Because it always zipped its lips."],
    ["Why did the squirrel sit near the fire?", "It wanted to be a little toastier."],
    ["What did one hiking boot say to the other?", "We make a great pair on the trail."]
  ];

  const jokeParas = [];
  jokes.forEach((j, i) => {
    jokeParas.push(new Paragraph({
      spacing: { before: 200, after: 40 },
      children: [new TextRun({ text: `${i + 1}. ${j[0]}`, size: 22, bold: true, color: "2C4A1E" })]
    }));
    jokeParas.push(new Paragraph({
      spacing: { after: 60 },
      indent: { left: 360 },
      children: [new TextRun({ text: j[1], size: 22, italics: true })]
    }));
  });

  return [
    ...activityTitle("Camping Jokes", "Read these around the picnic table, by the fire, or while waiting for the ferry."),
    ...jokeParas
  ];
}

// ===== PAGE 10: MY CAMPING MEMORIES =====
function memoriesPage() {
  const prompts = [
    "I went camping at",
    "The weather was",
    "I slept in a",
    "My favorite meal was",
    "The best activity was",
    "The funniest thing that happened was",
    "Next time I want to"
  ];

  const promptParas = [];
  prompts.forEach(p => {
    promptParas.push(new Paragraph({
      spacing: { before: 120, after: 20 },
      children: [new TextRun({ text: p + " ", size: 22, bold: true, color: "2C4A1E" })]
    }));
    promptParas.push(new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "BBBBBB", space: 1 } },
      spacing: { after: 80 },
      children: [new TextRun({ text: " ", size: 22 })]
    }));
  });

  return [
    ...activityTitle("My Camping Memories", "Your first camping trip is now part of your explorer story!"),
    new Paragraph({
      spacing: { after: 200 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "My First Camping Trip!", size: 28, bold: true })]
    }),
    ...promptParas,
    new Paragraph({ spacing: { before: 200 }, children: [] }),
    drawingBox("Draw your favorite camping memory", 7),
    new Paragraph({
      spacing: { before: 200 },
      children: [
        new TextRun({ text: "I rate this trip: ", size: 22, bold: true }),
        new TextRun({ text: "\u2606 \u2606 \u2606 \u2606 \u2606", size: 32 }),
        new TextRun({ text: "  (color in your stars!)", size: 18, italics: true, color: "888888" })
      ]
    })
  ];
}

// ===== BUILD DOCUMENT =====
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 24 } }
    }
  },
  sections: [
    // Cover / Title
    {
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "Orcas Island Activity Book", size: 16, italics: true, color: "888888" })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Page ", size: 16, color: "888888" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" })]
          })]
        })
      },
      children: [
        new Paragraph({ spacing: { before: 2000 }, children: [] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "First Camping Trip", size: 52, bold: true, font: "Georgia", color: "2C4A1E" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Activity Book", size: 44, bold: true, font: "Georgia", color: "2C4A1E" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "Moran State Park, Orcas Island, Washington", size: 24, italics: true })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "June 17-19, 2026", size: 24 })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 600 },
          children: [new TextRun({ text: "For one brave 6-year-old explorer", size: 22, italics: true, color: "555555" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "plus 2 happy camp dogs", size: 22, italics: true, color: "555555" })]
        }),
        // Page break then activities
        pageBreak(),
        ...scavengerHuntPage(),
        pageBreak(),
        ...bingoPage(),
        pageBreak(),
        ...wordSearchPage(),
        pageBreak(),
        ...journalPage(1, "Arrival Day", "First look at camp, first camp smells, first big adventure feeling."),
        pageBreak(),
        ...journalPage(2, "Adventure Day", "Waterfalls, swimming, campfire stories, and stars."),
        pageBreak(),
        ...journalPage(3, "Going Home", "Last lake visit, packing up, and ferry ride back."),
        pageBreak(),
        ...decoderPage(),
        pageBreak(),
        ...dotToDotPage(),
        pageBreak(),
        ...jokesPage(),
        pageBreak(),
        ...memoriesPage()
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "C:\\Users\\megas\\OneDrive - Microsoft\\ghcli-working\\epic-event-planner\\output\\orcas-island-camping\\printables\\Activity-Book.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Created: " + outPath);
});
