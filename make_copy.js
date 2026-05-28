const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, ExternalHyperlink, PageNumber, Header, Footer, PageBreak
} = require('./node_modules/docx');
const fs = require('fs');

// OpenSports URLs
const OS_ALL       = "https://opensports.net/sapphire-coast-pickleball/tab/events";
const OS_BERMAGUI  = "https://opensports.net/sapphire-coast-pickleball/tab/events?tags=Bermagui";
const OS_PAMBULA   = "https://opensports.net/sapphire-coast-pickleball/tab/events?tags=Pambula";
const OS_BEGA      = "https://opensports.net/sapphire-coast-pickleball/tab/events?tags=Bega";

const TEAL="006D77", GOLD="E8C547", LGREY="F2F2F2", DGREY="444444", WHITE="FFFFFF", PALE="F0F8F9";
const bd={style:BorderStyle.SINGLE,size:1,color:"DDDDDD"};
const borders={top:bd,bottom:bd,left:bd,right:bd};

// ── helpers ──────────────────────────────────────────────────────────────────
const r = (t,opts={}) => new TextRun({text:t,size:22,font:"Arial",color:DGREY,...opts});
const rb = t => r(t,{bold:true});

function h1(t) { return new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun({text:t,bold:true,size:40,font:"Arial",color:WHITE})],shading:{fill:TEAL,type:ShadingType.CLEAR},spacing:{before:400,after:200},indent:{left:160,right:160}}); }
function h2(t) { return new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun({text:t,bold:true,size:30,font:"Arial",color:TEAL})],spacing:{before:360,after:120},border:{bottom:{style:BorderStyle.SINGLE,size:6,color:TEAL,space:4}}}); }
function h3(t) { return new Paragraph({heading:HeadingLevel.HEADING_3,children:[new TextRun({text:t,bold:true,size:24,font:"Arial",color:DGREY})],spacing:{before:240,after:80}}); }
function h4(t) { return new Paragraph({children:[new TextRun({text:t,bold:true,size:22,font:"Arial",color:TEAL})],spacing:{before:180,after:60}}); }
function body(t) { return new Paragraph({children:[r(t)],spacing:{before:60,after:80}}); }
function sp() { return new Paragraph({children:[r("")],spacing:{before:40,after:40}}); }
function pb() { return new Paragraph({children:[new PageBreak()]}); }

function note(t) {
  return new Paragraph({children:[new TextRun({text:"SQUARESPACE: ",bold:true,size:19,font:"Arial",color:"777777"}),new TextRun({text:t,size:19,font:"Arial",color:"777777",italics:true})],shading:{fill:"F5F5F5",type:ShadingType.CLEAR},spacing:{before:60,after:60},indent:{left:160,right:160}});
}

function sectionLabel(t) {
  return new Paragraph({children:[new TextRun({text:t,bold:true,size:18,font:"Arial",color:WHITE})],shading:{fill:"AAAAAA",type:ShadingType.CLEAR},spacing:{before:200,after:0},indent:{left:120,right:120}});
}

function copyBlock(lines) {
  const kids = lines.map(l =>
    l === "" ? new Paragraph({children:[r("")],spacing:{before:20,after:20}})
             : new Paragraph({children:[r(l,{size:22,color:"111111"})],spacing:{before:40,after:40}})
  );
  return new Table({width:{size:9200,type:WidthType.DXA},columnWidths:[9200],rows:[new TableRow({children:[new TableCell({
    borders:{top:{style:BorderStyle.SINGLE,size:6,color:GOLD},bottom:{style:BorderStyle.SINGLE,size:1,color:"DDDDDD"},left:{style:BorderStyle.SINGLE,size:14,color:GOLD},right:{style:BorderStyle.SINGLE,size:1,color:"DDDDDD"}},
    width:{size:9200,type:WidthType.DXA},shading:{fill:"FFFEF0",type:ShadingType.CLEAR},
    margins:{top:120,bottom:120,left:220,right:120},children:kids
  })]})]}); 
}

function linkRow(label, url) {
  return new Paragraph({
    children:[
      new TextRun({text:"  LINK — "+label+": ",bold:true,size:20,font:"Arial",color:"555555"}),
      new ExternalHyperlink({link:url,children:[new TextRun({text:url,size:20,font:"Arial",color:"0070C0",underline:{}})]})
    ],
    spacing:{before:40,after:40},
    shading:{fill:"EEF6FB",type:ShadingType.CLEAR},
    indent:{left:160}
  });
}

function divider() {
  return new Paragraph({children:[r("")],border:{bottom:{style:BorderStyle.SINGLE,size:2,color:"DDDDDD",space:4}},spacing:{before:160,after:160}});
}


// ── PAGE CONTENT ─────────────────────────────────────────────────────────────

const homepageContent = [

  h2("PAGE 1 — HOMEPAGE"),
  note("Page type: Regular page. In Squarespace Fluid Engine, build each section below as a separate row."),
  sp(),

  // HERO
  sectionLabel("SECTION A — HERO BANNER"),
  note("Use a full-width Cover or Banner section. Set background to a pickleball action photo (replace the current Camel Rock image). Add two Button blocks side by side."),
  sp(),
  h4("Page Title (SEO / Browser Tab)"),
  copyBlock(["Sapphire Coast Pickleball — Play, Connect, Have Fun"]),
  sp(),
  h4("Meta Description (SEO tab)"),
  copyBlock(["Play pickleball on the Sapphire Coast NSW. 3 venues in Bermagui, Bega & Pambula. $10 per session, all equipment provided. Beginners welcome."]),
  sp(),
  h4("Hero Headline"),
  copyBlock(["Pickleball on the Sapphire Coast"]),
  sp(),
  h4("Hero Subheading"),
  copyBlock(["Join one of the fastest-growing sports in the world — right here in our backyard.", "All equipment provided. Just turn up."]),
  sp(),
  h4("Button 1 (Primary — filled style)"),
  copyBlock(["Label:  Book a Session"]),
  linkRow("Book a Session button — links to all SCPAI events on OpenSports", OS_ALL),
  sp(),
  h4("Button 2 (Secondary — outlined style)"),
  copyBlock(["Label:  Become a Member", "URL:    /register-join-us"]),
  divider(),

  // STATS
  sectionLabel("SECTION B — QUICK STATS BAR"),
  note("Add a narrow section with a dark background (use your teal colour). Place 4 text blocks side by side. Keep text large and bold."),
  sp(),
  copyBlock([
    "Stat 1:   3 Venues",
    "Stat 2:   $10 per Session",
    "Stat 3:   All Equipment Provided",
    "Stat 4:   Bermagui · Bega · Pambula",
  ]),
  divider(),

  // WHY
  sectionLabel("SECTION C — WHY PICKLEBALL? (3 columns)"),
  note("Add a 3-column row. Each column: Icon block at top, then a bold heading, then body text. Use your teal colour for icons."),
  sp(),
  h4("Column 1"),
  copyBlock(["Heading:  Easy to Learn","Body:     Pick up the basics in minutes. No racquet experience needed —","          we provide paddles, balls and all the gear."]),
  sp(),
  h4("Column 2"),
  copyBlock(["Heading:  Social & Fun","Body:     Meet your neighbours, make new friends, and share a laugh","          across the net. Pickleball is as much about community as sport."]),
  sp(),
  h4("Column 3"),
  copyBlock(["Heading:  Great for Every Age","Body:     Low-impact and high-enjoyment. Whether you're 25 or 75,","          pickleball is designed for every body."]),
  divider(),

  // NEW TO
  sectionLabel("SECTION D — NEW TO PICKLEBALL? (2 columns)"),
  note("2-column layout: left = text + link, right = photo (e.g. court diagram or fun action shot)."),
  sp(),
  h4("Heading"),
  copyBlock(["New to Pickleball?"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "Never picked up a paddle? No problem.",
    "",
    "You can play with us up to three times before deciding to become a member.",
    "Sessions cost just $10 and all gear is provided — just bring comfortable shoes",
    "and a smile.",
    "",
    "Come along, give it a go, and see why everyone is talking about it.",
  ]),
  sp(),
  h4("Text Link"),
  copyBlock(["Label:  See how sessions work →", "URL:    /venues  (or /find-a-session once you rename the page)"]),
  divider(),

  // VENUES
  sectionLabel("SECTION E — OUR VENUES (3 cards)"),
  note("3-column card layout. Each card: venue name as heading, details as body, one Book button. Use a light background on this section."),
  sp(),
  h4("Card 1 — Bermagui"),
  copyBlock(["Heading:    Bermagui","Subheading: Bermagui Indoor Stadium","Address:    20 Bunga Street, Bermagui NSW 2546","Type:       Indoors","Button:     Book at Bermagui →"]),
  linkRow("Book at Bermagui button", OS_BERMAGUI),
  sp(),
  h4("Card 2 — Pambula"),
  copyBlock(["Heading:    Pambula","Subheading: Pambula Sports Complex","Address:    Pambula Beach Road, Pambula NSW 2549","Type:       Outdoors","Button:     Book at Pambula →"]),
  linkRow("Book at Pambula button", OS_PAMBULA),
  sp(),
  h4("Card 3 — Bega"),
  copyBlock(["Heading:    Bega","Subheading: Sapphire Coast Anglican College Gymnasium","Address:    2 Max Slater Drive, Bega NSW 2550","Type:       Indoors","Button:     Book at Bega →"]),
  linkRow("Book at Bega button", OS_BEGA),
  divider(),

  // TESTIMONIALS
  sectionLabel("SECTION F — MEMBER TESTIMONIALS"),
  note("Use a Quote block, or a 2-column text layout with a coloured background. Replace the example quotes below with real ones from your members — ask at your next session!"),
  sp(),
  h4("Quote 1 (replace with real member quote)"),
  copyBlock(["\"I had never played before and was welcomed with open arms. Now I play twice a week!\"","— Sandra, Bermagui"]),
  sp(),
  h4("Quote 2 (replace with real member quote)"),
  copyBlock(["\"The best $10 I spend every week. Great exercise and even better company.\"","— Mike, Pambula"]),
  sp(),
  h4("Quote 3 (replace with real member quote)"),
  copyBlock(["\"I was nervous at first, but everyone was so encouraging. I was hooked after one session.\"","— Jenny, Bega"]),
  divider(),

  // NEWSLETTER
  sectionLabel("SECTION G — NEWSLETTER SIGNUP"),
  note("Use your existing Squarespace newsletter block — just update the heading and body text below."),
  sp(),
  copyBlock(["Heading:  Stay in the Loop","Body:     Get session updates, club news, and tips delivered straight to your inbox.","          No spam — just pickleball."]),
];


const sessionContent = [

  h2("PAGE 2 — FIND A SESSION"),
  note("Rename the current 'Venues' page to 'Find a Session'. Merge content from 'View & Book Sessions' into this page and delete the old page. URL slug: /find-a-session"),
  sp(),

  h4("Page Title (SEO / Browser Tab)"),
  copyBlock(["Find a Pickleball Session — Sapphire Coast"]),
  sp(),
  h4("Meta Description (SEO tab)"),
  copyBlock(["Find and book a pickleball session near you. Venues in Bermagui, Bega & Pambula. All levels welcome, 18 years and over, just $10 per session."]),
  divider(),

  sectionLabel("SECTION A — PAGE INTRO"),
  note("Full-width text section at the top of the page. Centre-aligned. Follow with a prominent button."),
  sp(),
  h4("Heading"),
  copyBlock(["Find a Session"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "We run sessions across three venues on the Sapphire Coast — from Bermagui to Pambula.",
    "",
    "All sessions are $10 per person for two hours of play. Equipment is provided — just",
    "bring yourself and comfortable shoes. Sessions are open to players 18 years and over.",
    "",
    "Book your spot on OpenSports. It's free to join our group.",
  ]),
  sp(),
  h4("Main Booking Button (large, prominent)"),
  copyBlock(["Label:  View All Sessions & Book Online"]),
  linkRow("Main booking button — all SCPAI events", OS_ALL),
  divider(),

  sectionLabel("SECTION B — VENUE CARDS"),
  note("3-column card layout, or stack vertically — one card per venue. Each card has a heading, details, and its own venue-filtered booking button."),
  sp(),

  h3("Venue 1 — Bermagui"),
  copyBlock([
    "Venue Name:   Bermagui Indoor Stadium",
    "Address:      20 Bunga Street, Bermagui NSW 2546",
    "Type:         Indoors — sheltered from weather",
    "Parking:      [Add parking details here]",
    "Sessions:     [Check OpenSports for current days and times]",
    "Note:         All sessions 18 years and over. All equipment provided.",
  ]),
  sp(),
  h4("Bermagui Booking Button"),
  copyBlock(["Label:  Book at Bermagui →"]),
  linkRow("Bermagui — filtered events", OS_BERMAGUI),
  sp(),

  h3("Venue 2 — Pambula"),
  copyBlock([
    "Venue Name:   Pambula Sports Complex",
    "Address:      Pambula Beach Road, Pambula NSW 2549",
    "Type:         Outdoors — please check conditions before attending",
    "Parking:      On-site parking available at the complex",
    "Sessions:     [Check OpenSports for current days and times]",
    "Note:         All sessions 18 years and over. All equipment provided.",
  ]),
  sp(),
  h4("Pambula Booking Button"),
  copyBlock(["Label:  Book at Pambula →"]),
  linkRow("Pambula — filtered events", OS_PAMBULA),
  sp(),

  h3("Venue 3 — Bega"),
  copyBlock([
    "Venue Name:   Sapphire Coast Anglican College Gymnasium",
    "Address:      2 Max Slater Drive, Bega NSW 2550",
    "Type:         Indoors — sheltered from weather",
    "Parking:      Follow the yellow route on the map to the car park.",
    "              Walk down to the gymnasium from the car park.",
    "Sessions:     [Check OpenSports for current days and times]",
    "Note:         All sessions 18 years and over. All equipment provided.",
  ]),
  sp(),
  h4("Bega Booking Button"),
  copyBlock(["Label:  Book at Bega →"]),
  linkRow("Bega — filtered events", OS_BEGA),
  divider(),

  sectionLabel("SECTION C — HOW IT WORKS (optional but recommended)"),
  note("A simple 3-step explainer helps first-timers feel confident. Use 3 icon columns or a numbered list."),
  sp(),
  h4("Heading"),
  copyBlock(["How It Works"]),
  sp(),
  copyBlock([
    "Step 1 — Find a Session",
    "Browse the sessions at your nearest venue and pick a time that suits you.",
    "",
    "Step 2 — Book Your Spot",
    "Click Book on OpenSports — it's free to join our group and takes just a minute.",
    "",
    "Step 3 — Turn Up & Play",
    "We provide paddles, balls, and a warm welcome. Just bring comfortable shoes.",
  ]),
  divider(),

  sectionLabel("SECTION D — VISITING THE SAPPHIRE COAST?"),
  note("Use a callout or coloured banner section near the bottom of the page."),
  sp(),
  h4("Heading"),
  copyBlock(["Visiting the Sapphire Coast?"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "Travellers and holidaymakers are very welcome to join a session.",
    "",
    "No membership required for your first three visits — just book online and show up.",
    "We'll look after you.",
  ]),
  sp(),
  h4("Button"),
  copyBlock(["Label:  Book a Session →"]),
  linkRow("Visitor booking button — all events", OS_ALL),
  divider(),

  sectionLabel("SECTION E — MEMBERSHIP NOTE"),
  note("Small text or a callout at the bottom of the page."),
  sp(),
  copyBlock([
    "After your third session, you'll need to register with Pickleball Association NSW",
    "(PANSW) and join SCPAI to continue playing at our venues.",
    "",
    "It's easy to do — head to the Join page to find out how.",
    "",
    "Link text:  How to become a member →   URL: /register-join-us",
  ]),
];


const aboutContent = [

  h2("PAGE 3 — ABOUT > OUR STORY"),
  note("Rename the current 'About SCPAI' page to 'Our Story'. URL slug: /our-story  (update any existing links)"),
  sp(),

  h4("Page Title (SEO / Browser Tab)"),
  copyBlock(["Our Story — Sapphire Coast Pickleball Association"]),
  sp(),
  h4("Meta Description (SEO tab)"),
  copyBlock(["Learn about the Sapphire Coast Pickleball Association — a volunteer-run, not-for-profit community club bringing people together across the NSW South Coast."]),
  divider(),

  sectionLabel("SECTION A — HERO TEXT"),
  note("Use a full-width banner section with your teal background colour and white text. No photo needed — strong typography works well here."),
  sp(),
  h4("Hero Headline"),
  copyBlock(["Bringing People Together Through the Joy of Pickleball"]),
  sp(),
  h4("Hero Subheading"),
  copyBlock(["A community-minded, not-for-profit association on the NSW Sapphire Coast."]),
  divider(),

  sectionLabel("SECTION B — MISSION"),
  note("Full-width text section, centred. Keep it short and punchy."),
  sp(),
  h4("Heading"),
  copyBlock(["Our Mission"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "At Sapphire Coast Pickleball Association Inc (SCPAI), our mission is simple:",
    "more courts, more players, more smiles.",
    "",
    "We believe pickleball is one of the best ways to stay active, meet people, and be",
    "part of a vibrant community. We want as many people as possible on the Sapphire",
    "Coast to experience that — whether they're a first-timer or a seasoned player.",
  ]),
  divider(),

  sectionLabel("SECTION C — WHO WE ARE"),
  note("2-column layout: left = text, right = group photo of members or committee (add your own photo here)."),
  sp(),
  h4("Heading"),
  copyBlock(["Who We Are"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "We are a group of volunteers who love pickleball and love the Sapphire Coast.",
    "",
    "SCPAI is a not-for-profit association that runs pickleball venues from Bermagui to",
    "Pambula. Our sessions are social, supportive, and a lot of fun — and everyone from",
    "beginners to experienced players is welcome.",
    "",
    "Whether you're a retiree looking for low-impact exercise, a competitive player",
    "wanting a challenge, or a complete beginner who just wants to try something new —",
    "you belong here.",
    "",
    "Please note: our sessions are currently open to players 18 years and over.",
  ]),
  divider(),

  sectionLabel("SECTION D — OUR COMMUNITY"),
  note("Full-width section. Works well with a light teal or pale background."),
  sp(),
  h4("Heading"),
  copyBlock(["Our Community"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "Pickleball is one of the fastest-growing sports in the world, and it's easy to see why.",
    "",
    "It brings people together across generations and fitness levels. The smaller court",
    "and lighter equipment make it accessible to almost everyone, while the social nature",
    "of the game means you'll leave every session with new friends — and probably a big smile.",
    "",
    "We've seen friendships form, fitness improve, and a real sense of community grow",
    "across all three of our venues. Come and be part of it.",
  ]),
  divider(),

  sectionLabel("SECTION E — GET INVOLVED"),
  note("Callout box or banner section at the bottom of the page."),
  sp(),
  h4("Heading"),
  copyBlock(["Get Involved"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "We are a volunteer-run organisation and we're always looking for enthusiastic people",
    "to help us grow pickleball on the Sapphire Coast.",
    "",
    "If you'd like to help run sessions, assist with venue management, or get involved",
    "with the committee, we'd love to hear from you.",
  ]),
  sp(),
  h4("Button"),
  copyBlock(["Label:  Get in Touch →","URL:    /contact"]),
];


const joinContent = [

  h2("PAGE 4 — JOIN > BECOME A MEMBER"),
  note("Rename the current 'Register' page to 'Become a Member'. URL slug: /become-a-member"),
  sp(),

  h4("Page Title (SEO / Browser Tab)"),
  copyBlock(["Become a Member — Sapphire Coast Pickleball Association"]),
  sp(),
  h4("Meta Description (SEO tab)"),
  copyBlock(["Join the Sapphire Coast Pickleball Association. Get unlimited sessions at 3 venues, PAA insurance cover, and access to a welcoming community. Easy 3-step sign-up."]),
  divider(),

  sectionLabel("SECTION A — HERO"),
  note("Banner section — teal background, white text. One CTA button."),
  sp(),
  h4("Headline"),
  copyBlock(["Ready to Join?"]),
  sp(),
  h4("Subheading"),
  copyBlock([
    "Membership gives you unlimited access to all SCPAI sessions, PAA insurance cover,",
    "and a growing community of players right here on the Sapphire Coast.",
  ]),
  divider(),

  sectionLabel("SECTION B — MEMBERSHIP BENEFITS (3 columns or icon list)"),
  note("Use an icon block row or a simple 3-column layout. This answers 'what do I get?' before asking them to sign up."),
  sp(),
  h4("Section Heading"),
  copyBlock(["What You Get as a Member"]),
  sp(),
  copyBlock([
    "Unlimited sessions at all 3 venues  —  Play as often as you like at Bermagui, Pambula and Bega.",
    "",
    "PAA insurance cover  —  Your Pickleball Association of Australia membership includes insurance.",
    "",
    "Community & connection  —  Be part of a welcoming group of players across the Sapphire Coast.",
    "",
    "Graded sessions  —  Play with others at a similar skill level as you develop your game.",
    "",
    "Member-only events  —  Access to club socials, tournaments, and special sessions.",
    "",
    "Club newsletter  —  Stay in the loop with news, tips, and session updates.",
    "",
    "Voting rights  —  Have your say at the Annual General Meeting.",
  ]),
  divider(),

  sectionLabel("SECTION C — HOW TO JOIN (3 steps)"),
  note("Use a numbered layout or 3 large icon columns — Step 1, Step 2, Step 3. This is the most important section on the page."),
  sp(),
  h4("Section Heading"),
  copyBlock(["How to Join — 3 Easy Steps"]),
  sp(),

  h3("Step 1 — Register with Pickleball Association NSW"),
  copyBlock([
    "First, register with Pickleball Association NSW (PANSW). This is done once",
    "and gives you your PAA member ID number and insurance cover.",
    "",
    "You'll complete this on the PANSW website — it only takes a few minutes.",
  ]),
  sp(),
  h4("Step 1 Button"),
  copyBlock(["Label:  Register with PANSW →","URL:    https://www.pansw.com.au  (check current PANSW registration URL)"]),
  sp(),

  h3("Step 2 — Join SCPAI"),
  copyBlock([
    "Once you have your PAA member ID, complete our club registration form to join",
    "the Sapphire Coast Pickleball Association.",
    "",
    "Annual membership fee: [INSERT CURRENT FEE — check with committee]",
  ]),
  sp(),
  h4("Step 2 Button"),
  copyBlock(["Label:  Join SCPAI →","URL:    [Keep your existing registration form link here]"]),
  sp(),

  h3("Step 3 — Book & Play!"),
  copyBlock([
    "You're all set! Book your next session on OpenSports, show your member ID",
    "number on arrival, and enjoy unlimited access to all three venues.",
  ]),
  sp(),
  h4("Step 3 Button"),
  copyBlock(["Label:  Book a Session on OpenSports →"]),
  linkRow("Step 3 button — all SCPAI sessions", OS_ALL),
  divider(),

  sectionLabel("SECTION D — ALREADY A MEMBER?"),
  note("A smaller callout box below the main steps. Helps existing members find what they need without cluttering the main flow."),
  sp(),
  h4("Heading"),
  copyBlock(["Already a Member?"]),
  sp(),
  h4("Body Text"),
  copyBlock([
    "If you're an existing SCPAI member renewing your membership, or need to update",
    "your details, use the link below.",
  ]),
  sp(),
  h4("Button"),
  copyBlock(["Label:  Existing Member  →","URL:    /register-join-us-1  (or update slug to /existing-member)"]),
  divider(),

  sectionLabel("SECTION E — FIND YOUR ID"),
  note("Small text section or a collapsible accordion item."),
  sp(),
  copyBlock([
    "Not sure of your PAA member ID number?",
    "",
    "Link text:  Find your PAA ID here →","URL:        /find-your-paa-id",
  ]),
  divider(),

  sectionLabel("SECTION F — QUESTIONS?"),
  sp(),
  copyBlock([
    "Have a question about membership? We're happy to help.",
    "",
    "Button label:  Contact Us  →","URL:           /contact",
  ]),
];


// ── ASSEMBLE DOC ──────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: { config: [{ reference:"bullets", levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] }] },
  styles: {
    default: { document: { run: { font:"Arial", size:22 } } },
    paragraphStyles: [
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:40,bold:true,font:"Arial",color:WHITE},paragraph:{spacing:{before:400,after:200},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:30,bold:true,font:"Arial",color:TEAL},paragraph:{spacing:{before:360,after:120},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:24,bold:true,font:"Arial",color:DGREY},paragraph:{spacing:{before:240,after:80},outlineLevel:2}},
    ],
  },
  sections: [{
    properties: { page: { size:{width:11906,height:16838}, margin:{top:1080,right:1080,bottom:1080,left:1080} } },
    headers: { default: new Header({ children:[new Paragraph({children:[new TextRun({text:"SCPAI — Website Copy  |  May 2026  |  Copy-paste ready content for Squarespace",size:18,font:"Arial",color:"888888"})],border:{bottom:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}},spacing:{after:80}})]})},
    footers: { default: new Footer({ children:[new Paragraph({children:[new TextRun({text:"Sapphire Coast Pickleball Association Inc  |  sapphirecoastpickleball.org.au    ",size:18,font:"Arial",color:"888888"}),new TextRun({text:"Page ",size:18,font:"Arial",color:"888888"}),new TextRun({children:[PageNumber.CURRENT],size:18,font:"Arial",color:"888888"})],border:{top:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}},alignment:AlignmentType.CENTER})]})},
    children: [

      // Cover
      new Paragraph({children:[new TextRun({text:"SAPPHIRE COAST PICKLEBALL ASSOCIATION INC",bold:true,size:48,font:"Arial",color:WHITE})],shading:{fill:TEAL,type:ShadingType.CLEAR},alignment:AlignmentType.CENTER,spacing:{before:600,after:0},indent:{left:200,right:200}}),
      new Paragraph({children:[new TextRun({text:"Website Copy — Ready to Paste into Squarespace",size:28,font:"Arial",color:GOLD})],shading:{fill:TEAL,type:ShadingType.CLEAR},alignment:AlignmentType.CENTER,spacing:{before:120,after:120},indent:{left:200,right:200}}),
      new Paragraph({children:[new TextRun({text:"4 Pages: Homepage · Find a Session · Our Story · Become a Member",size:22,font:"Arial",color:"CCEEEE"})],shading:{fill:TEAL,type:ShadingType.CLEAR},alignment:AlignmentType.CENTER,spacing:{before:0,after:500},indent:{left:200,right:200}}),

      new Paragraph({children:[new TextRun({text:"How to use this document",bold:true,size:24,font:"Arial",color:TEAL})],spacing:{before:200,after:80}}),
      new Paragraph({children:[new TextRun({text:"Each page is divided into sections. For every section you will find:",size:22,font:"Arial",color:DGREY})],spacing:{before:60,after:60}}),
      new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:"A grey SQUARESPACE note telling you which block type to use",size:22,font:"Arial",color:"777777",italics:true})],spacing:{before:40,after:40}}),
      new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:"Gold-bordered boxes with the exact copy to paste in",size:22,font:"Arial",color:DGREY})],spacing:{before:40,after:40}}),
      new Paragraph({numbering:{reference:"bullets",level:0},children:[new TextRun({text:"Blue LINK lines with the exact URL to use for each button",size:22,font:"Arial",color:DGREY})],spacing:{before:40,after:40}}),
      sp(),
      new Paragraph({children:[new TextRun({text:"OpenSports venue-filtered links used in this document:",bold:true,size:22,font:"Arial",color:TEAL})],spacing:{before:120,after:80}}),
      new Paragraph({children:[new TextRun({text:"All events:  ",bold:true,size:21,font:"Arial",color:DGREY}),new ExternalHyperlink({link:OS_ALL,children:[new TextRun({text:OS_ALL,size:21,font:"Arial",color:"0070C0",underline:{}})]})],spacing:{before:40,after:40}}),
      new Paragraph({children:[new TextRun({text:"Bermagui:    ",bold:true,size:21,font:"Arial",color:DGREY}),new ExternalHyperlink({link:OS_BERMAGUI,children:[new TextRun({text:OS_BERMAGUI,size:21,font:"Arial",color:"0070C0",underline:{}})]})],spacing:{before:40,after:40}}),
      new Paragraph({children:[new TextRun({text:"Pambula:     ",bold:true,size:21,font:"Arial",color:DGREY}),new ExternalHyperlink({link:OS_PAMBULA,children:[new TextRun({text:OS_PAMBULA,size:21,font:"Arial",color:"0070C0",underline:{}})]})],spacing:{before:40,after:40}}),
      new Paragraph({children:[new TextRun({text:"Bega:        ",bold:true,size:21,font:"Arial",color:DGREY}),new ExternalHyperlink({link:OS_BEGA,children:[new TextRun({text:OS_BEGA,size:21,font:"Arial",color:"0070C0",underline:{}})]})],spacing:{before:40,after:40}}),
      pb(),

      ...homepageContent, pb(),
      ...sessionContent,  pb(),
      ...aboutContent,    pb(),
      ...joinContent,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/sessions/confident-zealous-gauss/mnt/outputs/SCPAI_Website_Copy.docx', buf);
  console.log('Done.');
});
