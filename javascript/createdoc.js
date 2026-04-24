const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageBreak, TableOfContents,
  SimpleField, TabStopType, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ─── Colors ────────────────────────────────────────────────────────────────
const C = {
  primary: "1A56DB", secondary: "0E9F6E", accent: "F05252",
  dark: "1E2A3A", lightBlue: "EBF5FB", lightGreen: "F0FDF4",
  lightRed: "FEF2F2", lightYellow: "FFFBEB", lightGray: "F8F9FA",
  medGray: "E5E7EB", border: "CBD5E1", text: "1F2937",
  muted: "6B7280", white: "FFFFFF", yellow: "FEF3C7",
  yellowBorder: "F59E0B", orange: "FF6B2B", purple: "7C3AED",
  code: "1E293B", codeBorder: "334155",
};

// ─── Helpers ───────────────────────────────────────────────────────────────
const thin = (color = C.border) => ({ style: BorderStyle.SINGLE, size: 1, color });
const allB = (color = C.border) => ({ top: thin(color), bottom: thin(color), left: thin(color), right: thin(color) });
const noB = () => ({ top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} });
const leftB = (color, size=12) => ({ top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, left:{style:BorderStyle.SINGLE, size, color} });

const blank = (n=1) => Array.from({length:n}, () => new Paragraph({ children:[new TextRun("")], spacing:{before:60,after:60} }));
const pageBreak = () => new Paragraph({ children:[new PageBreak()] });

const p = (text, opts={}) => new Paragraph({
  children: [new TextRun({ text, font:"Arial", size:opts.size||22, bold:opts.bold, color:opts.color||C.text, italics:opts.italic })],
  alignment: opts.align||AlignmentType.LEFT,
  spacing: { before:opts.before??80, after:opts.after??80, line:opts.line||320 },
  ...(opts.indent ? {indent:{left:opts.indent}} : {}),
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font:"Arial", size:36, bold:true, color:C.primary })],
  spacing: {before:480, after:200},
  border: {bottom:{style:BorderStyle.SINGLE, size:6, color:C.primary, space:6}},
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font:"Arial", size:28, bold:true, color:C.dark })],
  spacing: {before:360, after:160},
});

const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text, font:"Arial", size:24, bold:true, color:C.primary })],
  spacing: {before:240, after:120},
});

const bullet = (text, bold=false, indent=720) => new Paragraph({
  numbering: { reference:"bullets", level:0 },
  children: [new TextRun({ text, font:"Arial", size:21, bold, color:C.text })],
  spacing: {before:60, after:60, line:300},
});

const numbered = (text) => new Paragraph({
  numbering: { reference:"numbers", level:0 },
  children: [new TextRun({ text, font:"Arial", size:21, color:C.text })],
  spacing: {before:60, after:60, line:300},
});

const codeBlock = (lines) => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:allB(C.codeBorder),
    shading:{fill:C.code,type:ShadingType.CLEAR},
    margins:{top:140,bottom:140,left:220,right:220},
    width:{size:9360,type:WidthType.DXA},
    children: lines.map(l => new Paragraph({
      children:[new TextRun({text:l||" ",font:"Courier New",size:18,color:"94D0FF"})],
      spacing:{before:20,after:20},
    }))
  })]})],
});

const infoBox = (icon, title, lines, bg=C.lightBlue, border=C.primary) => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders: leftB(border, 16),
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:140,bottom:140,left:200,right:200},
    width:{size:9360,type:WidthType.DXA},
    children:[
      new Paragraph({children:[
        new TextRun({text:icon+" ", font:"Arial", size:22}),
        new TextRun({text:title, font:"Arial", size:22, bold:true, color:border}),
      ], spacing:{before:40,after:80}}),
      ...lines.map(l => new Paragraph({
        children:[new TextRun({text:l, font:"Arial", size:20, color:C.text})],
        spacing:{before:30,after:30}, indent:{left:120},
      })),
    ]
  })]})],
});

const simpleTable = (headers, rows, colWidths) => new Table({
  width:{size:colWidths.reduce((a,b)=>a+b,0),type:WidthType.DXA},
  columnWidths: colWidths,
  rows:[
    new TableRow({tableHeader:true, children: headers.map((h,i) => new TableCell({
      borders:allB(C.primary), shading:{fill:C.primary,type:ShadingType.CLEAR},
      margins:{top:100,bottom:100,left:120,right:120}, width:{size:colWidths[i],type:WidthType.DXA},
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:h,font:"Arial",size:20,bold:true,color:C.white})]})]
    }))}),
    ...rows.map((row,ri) => new TableRow({children: row.map((cell,ci) => new TableCell({
      borders:allB(C.border), shading:{fill:ri%2===0?C.white:C.lightGray,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:120,right:120}, width:{size:colWidths[ci],type:WidthType.DXA},
      verticalAlign:VerticalAlign.CENTER,
      children:[new Paragraph({children:[new TextRun({text:cell,font:"Arial",size:20,color:C.text})]})]
    }))})),
  ],
});

const stepCard = (num, title, sub, color=C.primary) => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[640,8720],
  rows:[new TableRow({children:[
    new TableCell({
      borders:noB(), shading:{fill:color,type:ShadingType.CLEAR},
      margins:{top:120,bottom:120,left:100,right:100}, width:{size:640,type:WidthType.DXA},
      verticalAlign:VerticalAlign.CENTER,
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:String(num),font:"Arial",size:28,bold:true,color:C.white})]})]
    }),
    new TableCell({
      borders:{top:thin(C.border),bottom:thin(C.border),right:thin(C.border),left:noB().left},
      shading:{fill:C.lightGray,type:ShadingType.CLEAR},
      margins:{top:120,bottom:120,left:180,right:180}, width:{size:8720,type:WidthType.DXA},
      children:[
        new Paragraph({children:[new TextRun({text:title,font:"Arial",size:22,bold:true,color:color})], spacing:{before:20,after:50}}),
        new Paragraph({children:[new TextRun({text:sub,font:"Arial",size:19,color:C.muted})], spacing:{before:0,after:20}}),
      ]
    }),
  ]})]
});

const fileCard = (filename, location, editor, purpose, editWhat) => new Table({
  width:{size:9360,type:WidthType.DXA}, columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:allB(C.border), shading:{fill:C.white,type:ShadingType.CLEAR},
    margins:{top:120,bottom:120,left:160,right:160}, width:{size:9360,type:WidthType.DXA},
    children:[
      new Paragraph({children:[new TextRun({text:filename,font:"Courier New",size:22,bold:true,color:C.primary})], spacing:{before:20,after:60}}),
      new Paragraph({children:[new TextRun({text:"📁 Location: ",font:"Arial",size:19,bold:true,color:C.muted}), new TextRun({text:location,font:"Courier New",size:19,color:C.text})], spacing:{before:20,after:30}}),
      new Paragraph({children:[new TextRun({text:"✏️  Open with: ",font:"Arial",size:19,bold:true,color:C.muted}), new TextRun({text:editor,font:"Arial",size:19,color:C.text})], spacing:{before:20,after:30}}),
      new Paragraph({children:[new TextRun({text:"🎯 Purpose: ",font:"Arial",size:19,bold:true,color:C.muted}), new TextRun({text:purpose,font:"Arial",size:19,color:C.text})], spacing:{before:20,after:30}}),
      new Paragraph({children:[new TextRun({text:"🔧 Edit করো: ",font:"Arial",size:19,bold:true,color:C.secondary}), new TextRun({text:editWhat,font:"Arial",size:19,color:C.text})], spacing:{before:20,after:20}}),
    ]
  })]})],
});

// ═══════════════════════════════════════════════════════════════════════════
// DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [
      { reference:"bullets", levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] },
      { reference:"numbers", levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] },
    ]
  },
  styles: {
    default: { document:{ run:{font:"Arial",size:22,color:C.text} } },
    paragraphStyles: [
      { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{size:36,bold:true,font:"Arial",color:C.primary},
        paragraph:{spacing:{before:480,after:200},outlineLevel:0} },
      { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{size:28,bold:true,font:"Arial",color:C.dark},
        paragraph:{spacing:{before:360,after:160},outlineLevel:1} },
      { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{size:24,bold:true,font:"Arial",color:C.primary},
        paragraph:{spacing:{before:240,after:120},outlineLevel:2} },
    ]
  },
  sections:[{
    properties:{
      page:{
        size:{width:11906,height:16838},
        margin:{top:1440,right:1260,bottom:1440,left:1260}
      }
    },
    headers:{
      default: new Header({children:[
        new Paragraph({
          children:[
            new TextRun({text:"BlockVote — Developer Guide",font:"Arial",size:18,color:C.muted}),
            new TextRun({text:"\t",font:"Arial"}),
            new TextRun({text:"File & Edit Reference",font:"Arial",size:18,color:C.muted}),
          ],
          tabStops:[{type:TabStopType.RIGHT,position:9026}],
          border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.border,space:6}},
        })
      ]})
    },
    footers:{
      default: new Footer({children:[
        new Paragraph({
          children:[
            new TextRun({text:"BlockVote Project Guide  |  2025",font:"Arial",size:18,color:C.muted}),
            new TextRun({text:"\t",font:"Arial"}),
            new TextRun({text:"Page ",font:"Arial",size:18,color:C.muted}),
            new SimpleField("PAGE"),
          ],
          tabStops:[{type:TabStopType.RIGHT,position:9026}],
          border:{top:{style:BorderStyle.SINGLE,size:4,color:C.border,space:6}},
        })
      ]})
    },
    children:[

      // ═══════════ COVER ═══════════
      ...blank(3),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"DEVELOPER GUIDE",font:"Arial",size:22,bold:true,color:C.muted,})], spacing:{before:0,after:200}}),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"BlockVote",font:"Arial",size:80,bold:true,color:C.primary})], spacing:{before:0,after:80}}),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"ফাইল গাইড, ইনস্টলেশন, এডিট ও আপগ্রেড",font:"Arial",size:34,color:C.dark})], spacing:{before:0,after:300}}),
      new Table({
        width:{size:7200,type:WidthType.DXA}, columnWidths:[7200],
        rows:[new TableRow({children:[new TableCell({
          borders:allB(C.primary), shading:{fill:C.lightBlue,type:ShadingType.CLEAR},
          margins:{top:200,bottom:200,left:300,right:300}, width:{size:7200,type:WidthType.DXA},
          children:[
            new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"Project Stack",font:"Arial",size:20,bold:true,color:C.primary})], spacing:{before:60,after:100}}),
            new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"Ethereum  ·  Solidity  ·  Python Flask  ·  face_recognition",font:"Arial",size:22,color:C.text})], spacing:{before:0,after:60}}),
            new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"React Native  ·  Hardhat  ·  Ethers.js  ·  MetaMask",font:"Arial",size:22,color:C.text})], spacing:{before:0,after:60}}),
          ]
        })]})],
      }),
      ...blank(2),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"Web + Mobile App  |  Admin Approval System  |  Face Verified Voting",font:"Arial",size:20,color:C.muted})]}),
      pageBreak(),

      // ═══════════ TOC ═══════════
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"সূচিপত্র",font:"Arial",size:36,bold:true,color:C.primary})], spacing:{before:0,after:400}}),
      new TableOfContents("সূচিপত্র", {hyperlink:true, headingStyleRange:"1-3"}),
      pageBreak(),

      // ═══════════ CH1: OVERVIEW ═══════════
      h1("১. Project Overview"),

      h2("১.১  কতটি Part আছে এবং কীভাবে কাজ করে"),
      p("BlockVote project মোট ৪টি আলাদা Part নিয়ে তৈরি যেগুলো একসাথে কাজ করে। একটি বন্ধ থাকলে পুরো system কাজ করে না।"),
      ...blank(1),
      simpleTable(
        ["Part","Folder","কী করে","কে চালায়"],
        [
          ["Smart Contract","contracts/","ভোট ও voter data blockchain এ রাখে","Ethereum Network"],
          ["Python Backend","backend/","Face Recognition এবং Voter DB","তোমার PC (Terminal)"],
          ["Website","website/","Browser এ vote দেওয়া যায়","Browser এ open করো"],
          ["Mobile App","mobile/","Phone এ vote দেওয়া যায়","Expo Go App"],
        ],
        [2000,2000,3360,2000]
      ),
      ...blank(1),
      infoBox("🔄","Complete Flow:", [
        "Voter → Register request দেয় (name + face photo)",
        "Python backend → Face save করে (SQLite database)",
        "Blockchain → requestRegistration() call হয়",
        "Admin panel → 'Pending' দেখায়",
        "Admin → Approve করে (backend + blockchain উভয়ে update)",
        "Voter → Vote tab এ যায় → Face verify হয় → castVote() call → Permanent record",
      ], C.lightBlue, C.primary),

      pageBreak(),

      // ═══════════ CH2: INSTALL ═══════════
      h1("২. কী কী Install করতে হবে"),

      h2("২.১  সফটওয়্যার Install (একবারই করতে হবে)"),
      ...blank(1),
      simpleTable(
        ["Software","Download Link","কীসের জন্য","Check Command"],
        [
          ["Node.js v18+","nodejs.org","Blockchain + React Native","node --version"],
          ["Python 3.10+","python.org","Face Recognition backend","python --version"],
          ["Git","git-scm.com","Code version control","git --version"],
          ["VS Code","code.visualstudio.com","Code editor (recommended)","—"],
          ["Google Chrome","google.com/chrome","MetaMask extension এর জন্য","—"],
          ["MetaMask","metamask.io","Crypto wallet (Chrome extension)","—"],
          ["Expo Go App","Play Store / App Store","Mobile app test করতে","—"],
        ],
        [1800,2400,3000,2160]
      ),
      ...blank(1),

      h2("২.২  VS Code Extensions (Recommended)"),
      bullet("Solidity (by Nomic Foundation) — Smart contract syntax highlight"),
      bullet("Python — Python code support"),
      bullet("ES7+ React Snippets — React Native shortcuts"),
      bullet("Prettier — Code formatting"),
      ...blank(1),

      h2("২.৩  Node.js Packages Install"),
      codeBlock([
        "# Project root এ (একবারই করো):",
        "cd blockvote",
        "npm install",
        "",
        "# Mobile এর জন্য (আলাদা করো):",
        "cd mobile",
        "npm install",
        "",
        "# Expo CLI (global, একবারই):",
        "npm install -g expo-cli",
      ]),
      ...blank(1),

      h2("২.৪  Python Packages Install"),
      codeBlock([
        "cd backend",
        "",
        "# Virtual environment তৈরি করো:",
        "python -m venv venv",
        "",
        "# Activate করো:",
        "# Windows:",
        "venv\\Scripts\\activate",
        "# Mac/Linux:",
        "source venv/bin/activate",
        "",
        "# Packages install:",
        "pip install flask flask-cors numpy Pillow opencv-python",
        "",
        "# face_recognition (সবচেয়ে কঠিন part):",
        "pip install cmake",
        "pip install dlib",
        "pip install face-recognition",
        "",
        "# সমস্যা হলে Windows এ:",
        "pip install dlib-binary   # pre-compiled version",
      ]),
      ...blank(1),

      infoBox("⚠️","face_recognition Install সমস্যা হলে OS অনুযায়ী:", [
        "Windows:  pip install cmake  →  pip install dlib-binary  →  pip install face-recognition",
        "Mac (M1/M2):  brew install cmake  →  pip install dlib  →  pip install face-recognition",
        "Ubuntu/Linux:  sudo apt-get install cmake build-essential libopenblas-dev  →  pip install face-recognition",
      ], C.lightYellow, C.yellowBorder),

      pageBreak(),

      // ═══════════ CH3: FILES ═══════════
      h1("৩. প্রতিটি ফাইল — কোথায়, কীভাবে, কী এডিট করবে"),

      h2("৩.১  Smart Contract"),
      fileCard(
        "BlockVote.sol",
        "blockvote/contracts/BlockVote.sol",
        "VS Code (Solidity extension সহ)",
        "সব voting logic এখানে — voter approval, vote casting, results",
        "Election name বদলাতে: constructor এর string বদলাও | নতুন function যোগ করতে: Solidity শিখতে হবে | Deploy করলে ABI automatically update হয়"
      ),
      ...blank(1),
      codeBlock([
        "// Election name বদলাতে এখানে যাও:",
        "constructor(string memory _name, ...) {",
        '  // scripts/deploy.js এ "University Student Council..." বদলাও',
        "}",
        "",
        "// Voter approval threshold বদলাতে:",
        "// approveVoter() function এ logic add করতে পারো",
        "",
        "// Compile করো:",
        "npx hardhat compile",
      ]),
      ...blank(1),

      h2("৩.২  Python Backend"),
      fileCard(
        "app.py",
        "blockvote/backend/app.py",
        "VS Code বা যেকোনো text editor",
        "Face Recognition API — voter register, face verify, admin approval",
        "Face match threshold বদলাতে: tolerance=0.5 → কম হলে বেশি strict | Database: blockvote.db automatically তৈরি হয় | Port বদলাতে: app.run(port=5000) এর 5000 বদলাও"
      ),
      ...blank(1),
      codeBlock([
        "# Face match কঠিন করতে (0.0=perfect, 1.0=no match):",
        "match = face_recognition.compare_faces([stored], enc, tolerance=0.5)",
        "#                                                         ^^^^",
        "#                              0.4 = কঠিন | 0.6 = সহজ",
        "",
        "# Database file location:",
        "DB_FILE = 'blockvote.db'  # same folder এ তৈরি হয়",
        "",
        "# Backend চালাতে:",
        "python app.py",
        "# → Running on http://localhost:5000",
      ]),
      ...blank(1),

      h2("৩.৩  Website"),
      fileCard(
        "index.html",
        "blockvote/website/index.html",
        "VS Code দিয়ে edit, Chrome এ double-click করে open",
        "Web voter interface — registration, voting, admin panel",
        "⚠️ Deploy করার পর CONTRACT_ADDRESS এবং BACKEND URL বদলাও (মাত্র ২ লাইন)"
      ),
      ...blank(1),
      codeBlock([
        "// ফাইলের শুরুতে এই ২টি লাইন বদলাও:",
        "const CONTRACT_ADDRESS = '0xABC...';  // deploy এর পর পাবে",
        "const BACKEND = 'http://localhost:5000';  // production এ real URL",
        "",
        "// Candidate card এর color বদলাতে:",
        "// --blue, --green, --red variable গুলো :root এ আছে",
        "",
        "// Election নাম দেখাতে:",
        "// DEMO_DATA.candidates array এ candidates list আছে (demo mode)",
      ]),
      ...blank(1),

      h2("৩.৪  Mobile App Config (সবচেয়ে গুরুত্বপূর্ণ)"),
      fileCard(
        "config.js",
        "blockvote/mobile/src/config.js",
        "VS Code",
        "Mobile app এর সব settings এক জায়গায়",
        "Deploy করলে CONTRACT_ADDRESS বদলাও | BACKEND_URL এ তোমার PC এর IP দাও (localhost কাজ করবে না) | RPC_URL এ Alchemy এর Sepolia URL দাও"
      ),
      ...blank(1),
      codeBlock([
        "// এই ৩টি জিনিস বদলাতে হবে:",
        "export const CONFIG = {",
        "  CONTRACT_ADDRESS: '0xABC...',   // deploy করার পর",
        "",
        "  // PC এর IP address (192.168.X.X):5000",
        "  // ipconfig (Windows) বা ifconfig (Mac) দিয়ে IP দেখো",
        "  BACKEND_URL: 'http://192.168.1.100:5000',",
        "",
        "  // Alchemy.com থেকে Sepolia RPC URL",
        "  RPC_URL: 'https://eth-sepolia.g.alchemy.com/v2/YOUR_KEY',",
        "};",
      ]),
      ...blank(1),

      h2("৩.৫  Mobile App Screens (কোনটা কী করে)"),
      simpleTable(
        ["Screen File","কাজ","Edit করবে কখন"],
        [
          ["ConnectScreen.js","Wallet connect করে login","Login UI বদলাতে"],
          ["HomeScreen.js","Dashboard — stats, quick actions","নতুন stats বা info যোগ করতে"],
          ["RegisterScreen.js","Voter self-registration with face","Registration form fields যোগ করতে"],
          ["VoteScreen.js","Candidate list + face verify + vote","UI বদলাতে বা extra validation যোগ করতে"],
          ["StatusScreen.js","Voter নিজের approval status দেখে","Status messages বদলাতে"],
          ["AdminScreen.js","Admin approve/reject + election control","নতুন admin feature যোগ করতে"],
        ],
        [2600,3600,3160]
      ),

      pageBreak(),

      // ═══════════ CH4: RUN ═══════════
      h1("৪. কীভাবে চালাবে (প্রতিদিন)"),

      h2("৪.১  প্রতিদিন যে ৩টি Terminal খুলতে হবে"),
      ...blank(1),
      stepCard(1, "Terminal 1 — Local Blockchain", "npx hardhat node  (project root এ)  →  20টি test account পাবে", C.primary),
      ...blank(1),
      stepCard(2, "Terminal 2 — Python Backend", "cd backend  →  source venv/bin/activate  →  python app.py", C.secondary),
      ...blank(1),
      stepCard(3, "Terminal 3 — Mobile (optional)", "cd mobile  →  expo start  →  QR code scan করো Expo Go দিয়ে", C.orange),
      ...blank(1),
      p("Website: কোনো Terminal লাগবে না — website/index.html double-click করো।", {color:C.muted, italic:true}),
      ...blank(1),

      h2("৪.২  First Time Deploy (একবারই)"),
      codeBlock([
        "# Terminal 1 চালু থাকলে Terminal 2 তে:",
        "npx hardhat run scripts/deploy.js --network localhost",
        "",
        "# Output দেখবে:",
        "# ✅ Contract deployed to: 0x5FbDB2315678afecb367f032d93F642f64180aa3",
        "#                          ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^",
        "#                          এই address টি copy করো!",
        "",
        "# তারপর paste করো:",
        "# website/index.html  → const CONTRACT_ADDRESS = '0x5FbDB...'",
        "# mobile/src/config.js → CONTRACT_ADDRESS: '0x5FbDB...'",
      ]),
      ...blank(1),

      h2("৪.৩  MetaMask Setup (একবারই)"),
      numbered("Chrome এ MetaMask extension install করো"),
      numbered("New wallet তৈরি করো — 12-word seed phrase লিখে রাখো"),
      numbered("Settings → Advanced → Show Test Networks → ON"),
      numbered("Network dropdown → 'Localhost 8545' select করো"),
      numbered("npx hardhat node output থেকে Account #0 এর Private Key copy করো"),
      numbered("MetaMask → Account icon → Import Account → Private Key paste করো"),
      numbered("এই account টাই Admin হবে!"),

      pageBreak(),

      // ═══════════ CH5: EDIT & UPGRADE ═══════════
      h1("৫. কীভাবে Edit ও Upgrade করবে"),

      h2("৫.১  সহজ Upgrades (Solidity জানা লাগবে না)"),
      ...blank(1),
      infoBox("✏️","Website এর রঙ বদলাতে:", [
        "website/index.html খোলো → :root এর CSS variables বদলাও",
        "--blue: #3b82f6  →  যেকোনো hex color দাও",
        "--green: #22c55e  →  success color",
        "--bg: #060a12  →  background color",
      ], C.lightBlue, C.primary),
      ...blank(1),
      infoBox("✏️","Candidate যোগ করতে:", [
        "Option 1 (Easy): Admin panel থেকে UI দিয়েই candidate add করো",
        "Option 2 (Code): scripts/deploy.js এ addCandidate() calls যোগ করো",
        "Note: Voting শুরু হলে আর candidate add করা যাবে না",
      ], C.lightGreen, C.secondary),
      ...blank(1),
      infoBox("✏️","Face Recognition কঠিন/সহজ করতে:", [
        "backend/app.py খোলো",
        "tolerance=0.5 → এই value টি বদলাও",
        "0.4 = কঠিন (exact match চাই) | 0.6 = সহজ (similar হলেই হবে)",
      ], C.lightYellow, C.yellowBorder),
      ...blank(1),

      h2("৫.২  Smart Contract Upgrade করতে চাইলে"),
      infoBox("⚠️","গুরুত্বপূর্ণ নিয়ম:", [
        "Contract deploy করলে সেটা permanent — পরিবর্তন করা যায় না",
        "Upgrade = নতুন contract deploy করতে হবে + নতুন address সব জায়গায় বদলাতে হবে",
        "Test: npx hardhat test দিয়ে আগে test করো",
        "Local এ ঠিক হলে Sepolia তে deploy করো",
      ], C.lightRed, C.accent),
      ...blank(1),
      codeBlock([
        "// নতুন feature যোগ করার workflow:",
        "// 1. contracts/BlockVote.sol এ function যোগ করো",
        "// 2. test/Voting.test.js এ test লেখো",
        "// 3. npx hardhat test  →  সব pass হলে",
        "// 4. npx hardhat run scripts/deploy.js --network localhost",
        "// 5. নতুন address website + mobile config এ বসাও",
      ]),
      ...blank(1),

      h2("৫.৩  Mobile App এ নতুন Screen যোগ করতে"),
      codeBlock([
        "// 1. নতুন file তৈরি করো:",
        "//    mobile/src/screens/NewScreen.js",
        "",
        "// 2. App.js এ import করো:",
        "import NewScreen from './src/screens/NewScreen';",
        "",
        "// 3. Tab.Navigator এ যোগ করো:",
        "<Tab.Screen name='New' component={NewScreen}",
        "  options={{ tabBarLabel: 'New', tabBarIcon: ({color}) =>",
        "    <Text style={{fontSize:20,color}}>🆕</Text> }}/>",
      ]),
      ...blank(1),

      h2("৫.৪  Sepolia Testnet এ Deploy করতে"),
      stepCard(1, "Alchemy Account", "alchemy.com → Sign Up → Create App → Ethereum Sepolia → HTTPS URL copy করো", C.primary),
      ...blank(1),
      stepCard(2, "Free Test ETH নাও", "sepoliafaucet.com → তোমার MetaMask address paste → Send ETH", C.secondary),
      ...blank(1),
      stepCard(3, ".env File তৈরি করো", ".env.example কপি করো → rename করো .env → values বসাও", C.orange),
      ...blank(1),
      stepCard(4, "Deploy করো", "npx hardhat run scripts/deploy.js --network sepolia → address note করো", C.purple),
      ...blank(1),
      codeBlock([
        "# .env file এ:",
        "SEPOLIA_RPC_URL=https://eth-sepolia.g.alchemy.com/v2/YOUR_KEY",
        "PRIVATE_KEY=your_metamask_private_key_here",
        "",
        "# Deploy:",
        "npx hardhat run scripts/deploy.js --network sepolia",
        "",
        "# ⚠️ .env কখনো GitHub এ push করো না!",
        "# .gitignore এ .env already আছে",
      ]),

      pageBreak(),

      // ═══════════ CH6: LEARN ═══════════
      h1("৬. কী কী শিখতে হবে"),

      h2("৬.১  Person A — Blockchain Developer"),
      ...blank(1),
      simpleTable(
        ["Topic","Platform","কত সময়","Priority"],
        [
          ["Solidity Basics","cryptozombies.io (free, game-based)","1-2 সপ্তাহ","🔴 Must"],
          ["Hardhat Framework","hardhat.org/docs","3-4 দিন","🔴 Must"],
          ["Ethers.js Basics","docs.ethers.org","3-4 দিন","🔴 Must"],
          ["Smart Contract Security","solidity-by-example.org","1 সপ্তাহ","🟡 Important"],
          ["Blockchain Fundamentals","YouTube: Patrick Collins","2-3 দিন","🟡 Important"],
          ["OpenZeppelin Contracts","docs.openzeppelin.com","পরে শিখো","🟢 Optional"],
        ],
        [2500,3000,1660,2200]
      ),
      ...blank(1),

      h2("৬.২  Person B — Backend + Frontend Developer"),
      ...blank(1),
      simpleTable(
        ["Topic","Platform","কত সময়","Priority"],
        [
          ["Python Flask","flask.palletsprojects.com/tutorial","3-4 দিন","🔴 Must"],
          ["face_recognition Library","github.com/ageitgey/face_recognition","2-3 দিন","🔴 Must"],
          ["React Native Basics","reactnative.dev/docs/getting-started","1-2 সপ্তাহ","🔴 Must"],
          ["Expo Framework","docs.expo.dev","3-4 দিন","🔴 Must"],
          ["React Navigation","reactnavigation.org/docs","2-3 দিন","🟡 Important"],
          ["Async/Await JavaScript","javascript.info/async-await","2 দিন","🟡 Important"],
        ],
        [2500,3400,1560,2100]
      ),
      ...blank(1),

      h2("৬.৩  দুজনের জন্য Common Knowledge"),
      bullet("Git & GitHub — code share করতে (gitpro.tpope.io — free book)", true),
      bullet("REST API — HTTP requests কীভাবে কাজ করে"),
      bullet("JSON — data format বোঝা"),
      bullet("Terminal/Command Line — basic commands"),
      bullet("MetaMask usage — wallet operations"),
      ...blank(1),
      infoBox("💡","Learning Order (Recommended):", [
        "Week 1: Git + Terminal basics + MetaMask setup",
        "Week 2: Person A → CryptoZombies Solidity | Person B → Flask + face_recognition",
        "Week 3: Person A → Hardhat + Ethers.js | Person B → React Native basics",
        "Week 4: Integration — দুজন মিলে connect করো + test করো",
      ], C.lightBlue, C.primary),

      pageBreak(),

      // ═══════════ CH7: QUICK REF ═══════════
      h1("৭. Quick Reference Card"),

      h2("৭.১  সব Important Commands"),
      codeBlock([
        "# ── Blockchain ──────────────────────────────────",
        "npx hardhat compile          # Contract compile করো",
        "npx hardhat test             # Tests চালাও",
        "npx hardhat node             # Local blockchain start",
        "npx hardhat run scripts/deploy.js --network localhost",
        "npx hardhat run scripts/deploy.js --network sepolia",
        "",
        "# ── Python Backend ──────────────────────────────",
        "cd backend",
        "venv\\Scripts\\activate       # Windows",
        "source venv/bin/activate      # Mac/Linux",
        "python app.py                 # Backend start (port 5000)",
        "",
        "# ── Mobile ──────────────────────────────────────",
        "cd mobile",
        "expo start                    # QR code দেখাবে",
        "expo start --android          # Android emulator",
        "",
        "# ── Test Backend ────────────────────────────────",
        "curl http://localhost:5000/health",
        "curl http://localhost:5000/admin/voters",
      ]),
      ...blank(1),

      h2("৭.২  Deploy করার পর ৩টি জায়গায় Address বসাতে হবে"),
      ...blank(1),
      simpleTable(
        ["File","কোথায়","কী বসাবে"],
        [
          ["website/index.html","const CONTRACT_ADDRESS = '0x...'","Deployed contract address"],
          ["mobile/src/config.js","CONTRACT_ADDRESS: '0x...'","Deployed contract address"],
          ["mobile/src/config.js","BACKEND_URL: 'http://IP:5000'","PC এর local IP address"],
        ],
        [3200,3360,2800]
      ),
      ...blank(1),

      h2("৭.৩  Common Errors ও সমাধান"),
      simpleTable(
        ["Error","কারণ","সমাধান"],
        [
          ["MetaMask: Wrong Network","localhost এ নেই","Network → Localhost 8545 select করো"],
          ["Backend 404","URL ভুল","http://localhost:5000/health চেক করো"],
          ["Face not detected","আলো কম","ভালো আলোতে সরাসরি তাকাও"],
          ["dlib install failed","CMake নেই","pip install cmake আগে করো"],
          ["Mobile can't connect","localhost কাজ না","PC এর IP দাও (192.168.X.X)"],
          ["Transaction reverted","Gas বা wrong account","Hardhat test account import করো"],
          ["Contract address wrong","পুরনো address","Re-deploy করো → নতুন address বসাও"],
        ],
        [2500,2600,4260]
      ),

      ...blank(2),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"— সমাপ্ত —",font:"Arial",size:22,color:C.muted,italics:true})], spacing:{before:200,after:100}}),
      new Paragraph({alignment:AlignmentType.CENTER, children:[new TextRun({text:"BlockVote Developer Guide  |  Academic Project 2025",font:"Arial",size:18,color:C.muted})]}),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/BlockVote-Developer-Guide.docx', buf);
  console.log('Done!');
}).catch(e => console.error(e));