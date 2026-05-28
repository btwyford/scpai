const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, ExternalHyperlink, PageNumber, Header, Footer, PageBreak
} = require('./node_modules/docx');
const fs = require('fs');

const OPENSPORTS_URL = "https://opensports.net/sapphire-coast-pickleball/tab/events";
const TEAL="006D77", GOLD="E8C547", LGREY="F2F2F2", DGREY="595959", WHITE="FFFFFF";
const bd = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text:t, bold:true, size:36, font:"Arial", color:WHITE })], shading:{fill:TEAL,type:ShadingType.CLEAR}, spacing:{before:360,after:180}, indent:{left:200,right:200} }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text:t, bold:true, size:28, font:"Arial", color:TEAL })], spacing:{before:320,after:120}, border:{bottom:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}} }); }
function h3(t) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text:t, bold:true, size:24, font:"Arial", color:DGREY })], spacing:{before:240,after:80} }); }
function body(t) { return new Paragraph({ children:[new TextRun({text:t,size:22,font:"Arial",color:DGREY})], spacing:{before:60,after:60} }); }
function run(t) { return new TextRun({text:t,size:22,font:"Arial",color:DGREY}); }
function bold(t) { return new TextRun({text:t,bold:true,size:22,font:"Arial",color:DGREY}); }
function bullet(t) { return new Paragraph({ numbering:{reference:"bullets",level:0}, children:[new TextRun({text:t,size:22,font:"Arial",color:DGREY})], spacing:{before:40,after:40} }); }
function bulletBold(l,r) { return new Paragraph({ numbering:{reference:"bullets",level:0}, children:[new TextRun({text:l,bold:true,size:22,font:"Arial",color:DGREY}),new TextRun({text:r,size:22,font:"Arial",color:DGREY})], spacing:{before:40,after:40} }); }
function nb(t) { return new Paragraph({ children:[new TextRun({text:"Note: ",bold:true,size:22,font:"Arial",color:TEAL}),new TextRun({text:t,size:22,font:"Arial",color:DGREY})], shading:{fill:"E8F4F5",type:ShadingType.CLEAR}, spacing:{before:80,after:80}, indent:{left:200,right:200} }); }
function linkPara(l,url) { return new Paragraph({ children:[new TextRun({text:l+": ",bold:true,size:22,font:"Arial",color:DGREY}), new ExternalHyperlink({link:url,children:[new TextRun({text:url,size:22,font:"Arial",color:"0070C0",underline:{}})]})] , spacing:{before:60,after:60} }); }
function pb() { return new Paragraph({children:[new PageBreak()]}); }
function sp() { return new Paragraph({children:[new TextRun("")],spacing:{before:60,after:60}}); }

function twoCell(a,b,shade) {
  const f=shade?LGREY:WHITE;
  return new TableRow({children:[
    new TableCell({borders,width:{size:3000,type:WidthType.DXA},shading:{fill:f,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[bold(a)]})]}),
    new TableCell({borders,width:{size:6360,type:WidthType.DXA},shading:{fill:f,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[run(b)]})]}),
  ]});
}
function twoCol(rows) { return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3000,6360],rows:rows.map((r,i)=>twoCell(r[0],r[1],i%2===0))}); }

function navHdr() {
  const c=(t,w)=>new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:{fill:TEAL,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:t,bold:true,size:22,font:"Arial",color:WHITE})]})]});
  return new TableRow({children:[c("CURRENT",2800),c("SUGGESTED",2800),c("REASON",3760)]});
}
function navRow(a,b,c,shade) {
  const f=shade?LGREY:WHITE;
  const cl=(t,w)=>new TableCell({borders,width:{size:w,type:WidthType.DXA},shading:{fill:f,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[run(t)]})]});
  return new TableRow({children:[cl(a,2800),cl(b,2800),cl(c,3760)]});
}
function navTable(rows) { return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2800,2800,3760],rows:[navHdr(),...rows.map((r,i)=>navRow(r[0],r[1],r[2],i%2===0))]}); }

function copyBox(label,lines) {
  const kids=[new Paragraph({children:[new TextRun({text:"Suggested copy: "+label,bold:true,size:20,font:"Arial",color:TEAL})],spacing:{before:80,after:60}}),...lines.map(l=>new Paragraph({children:[new TextRun({text:l,size:21,font:"Arial",color:"333333",italics:true})],spacing:{before:40,after:40}}))];
  return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({borders:{top:{style:BorderStyle.SINGLE,size:4,color:GOLD},bottom:{style:BorderStyle.SINGLE,size:1,color:"DDDDDD"},left:{style:BorderStyle.SINGLE,size:12,color:GOLD},right:{style:BorderStyle.SINGLE,size:1,color:"DDDDDD"}},width:{size:9360,type:WidthType.DXA},shading:{fill:"FFFDF0",type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:200,right:120},children:kids})]})]}); 
}


const doc = new Document({
  numbering: { config: [{ reference:"bullets", levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}] }] },
  styles: {
    default: { document: { run: { font:"Arial", size:22 } } },
    paragraphStyles: [
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:36,bold:true,font:"Arial",color:WHITE},paragraph:{spacing:{before:360,after:180},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:28,bold:true,font:"Arial",color:TEAL},paragraph:{spacing:{before:320,after:120},outlineLevel:1}},
      {id:"Heading3",name:"Heading 3",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:24,bold:true,font:"Arial",color:DGREY},paragraph:{spacing:{before:240,after:80},outlineLevel:2}},
    ],
  },
  sections: [{
    properties: { page: { size:{width:11906,height:16838}, margin:{top:1080,right:1080,bottom:1080,left:1080} } },
    headers: { default: new Header({ children: [new Paragraph({ children:[new TextRun({text:"SCPAI Website Revamp — Content Brief  |  May 2026",size:18,font:"Arial",color:"888888"})], border:{bottom:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}}, spacing:{after:80} })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ children:[new TextRun({text:"Sapphire Coast Pickleball Association Inc  |  sapphirecoastpickleball.org.au    ",size:18,font:"Arial",color:"888888"}),new TextRun({text:"Page ",size:18,font:"Arial",color:"888888"}),new TextRun({children:[PageNumber.CURRENT],size:18,font:"Arial",color:"888888"})], border:{top:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}}, alignment:AlignmentType.CENTER })] }) },
    children: [
      // COVER
      new Paragraph({children:[new TextRun({text:"SAPPHIRE COAST PICKLEBALL ASSOCIATION INC",bold:true,size:48,font:"Arial",color:WHITE})],shading:{fill:TEAL,type:ShadingType.CLEAR},alignment:AlignmentType.CENTER,spacing:{before:600,after:0},indent:{left:200,right:200}}),
      new Paragraph({children:[new TextRun({text:"Website Revamp — Full Content Brief",size:30,font:"Arial",color:GOLD})],shading:{fill:TEAL,type:ShadingType.CLEAR},alignment:AlignmentType.CENTER,spacing:{before:120,after:500},indent:{left:200,right:200}}),
      twoCol([
        ["Prepared for","Barnaby Twyford / SCPAI Committee"],
        ["Date","May 2026"],
        ["Platform","Squarespace (existing site)"],
        ["Current URL","sapphirecoastpickleball.org.au"],
        ["Booking system","OpenSports — see link throughout this document"],
        ["Scope","Full site — all pages + navigation structure"],
        ["Goals","Attract new players  |  Serve members  |  Modern professional look"],
      ]),
      sp(),
      nb("This document is a content and structure guide only. All suggested copy is a starting point — edit to match your club's voice."),
      pb(),

      // S1
      h1("1.  What the Current Site Does Well — and Where It Struggles"),
      sp(),h2("1.1  Strengths to Keep"),
      bullet("Already on Squarespace — no platform migration needed."),
      bullet("Beautiful Camel Rock hero photo establishes a strong sense of place."),
      bullet("Good foundational content exists (What is Pickleball?, Gradings, Rules)."),
      bullet("OpenSports integration for bookings is working."),
      bullet("Clear not-for-profit community ethos comes through."),
      sp(),h2("1.2  Key Issues Found"),
      bulletBold("Cluttered navigation: ","The Play dropdown has 7 items and The Association is buried. Visitors cannot quickly find How do I join? or When can I play?."),
      bulletBold("Duplicate Insurance page: ","Insurance appears under both Membership and The Association — confusing and looks like an error."),
      bulletBold("Weak homepage: ","The hero photo is a scenic landscape, not a pickleball action shot. There are no clear calls-to-action or quick stats to hook first-time visitors."),
      bulletBold("No News or Blog section: ","There is nowhere to post tournament results, club milestones, or tips."),
      bulletBold("No photo gallery: ","Social sport clubs thrive on showing happy people. There are currently no member photos."),
      bulletBold("Meta titles need fixing: ","The What is Pickleball? page is titled About 1 in the browser tab and social share cards."),
      bulletBold("18+ restriction not prominent: ","Mentioned only in small print. New visitors booking through OpenSports may arrive unaware."),
      bulletBold("No testimonials or member voices: ","Nothing that shows the community warmth of the club."),
      pb(),

      // S2
      h1("2.  Navigation Restructure"),
      sp(),
      body("Squarespace supports up to 5-6 top-level nav items plus dropdowns. The suggested structure gets any visitor to the right page in one or two clicks."),
      sp(),h2("2.1  Proposed Top Navigation"),
      bullet("Home"),bullet("About (dropdown)"),bullet("Play (dropdown)"),bullet("Join (dropdown)"),bullet("News"),bullet("Contact"),
      sp(),h2("2.2  Dropdown Contents"),
      h3("About"),bullet("Our Story"),bullet("What is Pickleball?"),bullet("Our Committee"),bullet("Sponsors"),
      sp(),h3("Play"),bullet("Find a Session  (Venues + booking on one page)"),bullet("Skill Levels & Gradings"),bullet("Rules"),bullet("Safety, Etiquette & Code of Conduct  (merged)"),bullet("FAQs"),
      sp(),h3("Join"),bullet("Become a Member"),bullet("Existing Member Login"),bullet("Member Benefits"),bullet("Insurance & Cover"),bullet("Find Your ID"),
      sp(),h2("2.3  What Changes and Why"),
      navTable([
        ["About Us","About","Shorter, cleaner label"],
        ["About SCPAI","Our Story","Friendlier — SCPAI acronym is unfamiliar to newcomers"],
        ["Play (7 items)","Play (5 items)","Merged Safety+Etiquette+Code of Conduct; merged Venues+Booking"],
        ["Membership","Join","Action-oriented; clearer for newcomers"],
        ["Register","Become a Member","More welcoming call-to-action language"],
        ["The Association","Moved into About","Consolidates governance; too formal for a community club"],
        ["Insurance (x2)","Insurance (x1 in Join)","Removes confusing duplication"],
        ["Contact Us (top)","Contact (top)","Consistent, shorter label"],
        ["AGM - 2025","About > Documents","Annual docs belong in governance section"],
        ["No News page","News  (new)","Blog/news for updates, results, tips"],
      ]),
      sp(),nb("In Squarespace: go to Pages > Navigation, drag pages into folders, rename items. Use Not Linked pages for dropdown folder labels."),
      pb(),

      // S3
      h1("3.  Page-by-Page Content Suggestions"),
      sp(),h2("3.1  Homepage"),
      body("The homepage needs to answer three questions within 5 seconds: What is this? Is it for me? What do I do next?"),
      sp(),h3("Section A — Hero (full-width banner)"),
      body("Replace the Camel Rock scenic photo with an action shot of people playing pickleball — smiling, mid-game. If you do not have one yet, save the Camel Rock image for the About page."),
      sp(),
      copyBox("Hero headline + subheading + CTAs",[
        "Headline:   Pickleball on the Sapphire Coast",
        "Subheading: Join one of the fastest-growing sports in the world — right here in our backyard.",
        "            All equipment provided. Just turn up.",
        "Button 1 (primary):   Book a Session  ->  OpenSports link (see below)",
        "Button 2 (secondary): Become a Member ->  /join",
      ]),
      linkPara("Book a Session button — link to OpenSports",OPENSPORTS_URL),
      sp(),h3("Section B — Quick Stats Bar"),
      copyBox("Stats bar",["3 Venues  |  $10 per Session  |  All Equipment Provided  |  Bermagui  Bega  Pambula"]),
      sp(),h3("Section C — Why Pickleball? (3-column feature row)"),
      copyBox("Three feature columns",[
        "Easy to Learn  —  Pick up the basics in minutes. No racquet experience needed.",
        "Social & Fun   —  Meet your neighbours, make new friends, share a laugh.",
        "Great for Every Age  —  Low-impact, high-enjoyment. Perfect from 25 to 75.",
      ]),
      sp(),h3("Section D — New to Pickleball?"),
      copyBox("New to Pickleball? block",[
        "Heading: New to Pickleball?",
        "Body:    Never picked up a paddle? No problem. You can play with us up to three times",
        "         before deciding to become a member. Just $10 and all gear is provided.",
        "Link:    Learn how sessions work -> /find-a-session",
      ]),
      sp(),h3("Section E — Our Venues"),
      body("Three venue cards in a row. Each shows venue name, suburb, indoor/outdoor, and a Book Now button linking to OpenSports."),
      linkPara("Book Now buttons on venue cards — link to",OPENSPORTS_URL),
      sp(),h3("Section F — Member Testimonials"),
      copyBox("Example quotes (replace with real member quotes)",[
        "\"I had never played before and was welcomed with open arms. Now I play twice a week!\"",
        "— Sandra, Bermagui",
        "",
        "\"The best $10 I spend every week. Great exercise and even better company.\"",
        "— Mike, Pambula",
      ]),
      sp(),h3("Section G — Sponsors Strip"),
      body("Horizontal logo strip near the bottom — logos only, linking to sponsor websites."),
      sp(),h3("Section H — Newsletter Signup"),
      copyBox("Newsletter copy",["Heading: Stay in the Loop","Body: Session updates, news, and tips — no spam, just pickleball."]),
      pb(),

      h2("3.2  About > Our Story"),
      copyBox("Page structure",[
        "Hero:    Bringing people together through the joy of pickleball.",
        "Mission: We are a volunteer-run, not-for-profit association passionate about",
        "         growing pickleball across the Sapphire Coast — from Bermagui to Pambula.",
        "Story:   [Keep existing text; add: when club was founded, member count, community built.]",
        "CTA:     Interested in volunteering? We'd love to hear from you. -> /contact",
      ]),
      sp(),nb("Add a group photo of members or committee — it makes the page feel real and trustworthy."),
      sp(),
      h2("3.3  About > What is Pickleball?"),
      body("The existing content is excellent. Main improvements are formatting and visual elements."),
      bullet("Add a short intro paragraph at the very top."),
      bullet("Break the history into a visually distinct callout or timeline."),
      bullet("Add a court diagram image (available from Pickleball Australia — check licensing)."),
      bullet("Add a Getting Started in 3 Steps section at the bottom linking to Find a Session."),
      bullet("Fix the page SEO title — currently shows as 'About 1'. Go to Pages > gear icon > SEO tab > Browser Tab."),
      copyBox("Getting Started in 3 Steps",[
        "1. Book a Session  ->  Head to Find a Session and pick a venue near you.",
        "2. Turn Up & Play  ->  We provide paddles and balls. Just wear comfortable shoes.",
        "3. Join the Club   ->  After your third session, become a member and play unlimited.",
      ]),
      pb(),

      h2("3.4  Play > Find a Session"),
      body("Merge the current Venues and View & Book Sessions pages into one well-organised page."),
      sp(),h3("Page layout"),
      bullet("Top: Bold intro + direct button to OpenSports booking."),
      bullet("Middle: Three venue cards with address, type, session days/times, and a Book button each."),
      bullet("Bottom: Visiting the Sapphire Coast? callout for tourists."),
      sp(),
      copyBox("Page intro",[
        "Heading: Find a Session",
        "Body:    We run sessions at three venues across the Sapphire Coast.",
        "         All sessions are $10 per person for two hours of play.",
        "         Equipment is provided. Sessions are open to players 18 years and over.",
        "Button:  View & Book Sessions on OpenSports  ->  link below",
      ]),
      linkPara("OpenSports — all SCPAI events (main booking button)",OPENSPORTS_URL),
      sp(),
      copyBox("Venue card details",[
        "BERMAGUI:  Bermagui Indoor Stadium, 20 Bunga Street — Indoors",
        "PAMBULA:   Pambula Sports Complex, Pambula Beach Road — Outdoors",
        "BEGA:      Sapphire Coast Anglican College Gymnasium, 2 Max Slater Dr — Indoors",
        "           Parking: Drive along yellow route on map; park and walk down to gym.",
        "",
        "Each card: Book a spot button -> OpenSports link below",
      ]),
      linkPara("OpenSports — use for all venue Book buttons",OPENSPORTS_URL),
      sp(),nb("OpenSports does not currently support venue-filtered deep links. Link all venue Book buttons to the main events page above. Update individually if OpenSports adds venue filters in future."),
      sp(),
      copyBox("Visiting the Sapphire Coast? callout",[
        "Heading: Visiting the Sapphire Coast?",
        "Body:    Travellers and holidaymakers are very welcome — no membership required",
        "         for your first three visits. Book online and show up.",
        "Button:  Book a Session  ->  link below",
      ]),
      linkPara("OpenSports — visitor booking link",OPENSPORTS_URL),
      sp(),

      h2("3.5  Play > Safety, Etiquette & Code of Conduct"),
      body("Merge the three current pages into one page with clear section headings. Use Squarespace's Accordion block to keep it tidy."),
      bullet("Section 1: On-Court Safety"),bullet("Section 2: Player Etiquette"),bullet("Section 3: Code of Conduct"),bullet("Section 4: Reporting a Concern  ->  /contact"),
      sp(),

      h2("3.6  Play > FAQs"),
      body("Use Squarespace's Accordion block. Suggested question groups:"),
      bulletBold("Getting Started: ","How much does it cost? Do I need equipment? Can I come as a beginner?"),
      bulletBold("Sessions & Booking: ","How do I book? What if it rains? Can I bring a friend?"),
      bulletBold("Membership: ","Do I have to be a member? How do I join? What does membership include?"),
      bulletBold("Rules: ","Where are the rules? What is the Kitchen? What skill level do I need?"),
      bulletBold("18+ Policy: ","Why is the club 18+? Are there junior clubs nearby?"),
      sp(),nb("Make the 18+ requirement a dedicated FAQ question to prevent confusion on arrival."),
      pb(),

      h2("3.7  Join > Become a Member"),
      copyBox("Page structure",[
        "Heading: Ready to Join?",
        "Subhead: Membership gives you unlimited sessions, PAA insurance, and a great community.",
        "",
        "Step 1 — Register with Pickleball Association NSW (PANSW)",
        "         [Button: Register with PANSW -> PANSW website]",
        "         Done once — gives you your PAA member ID and insurance cover.",
        "",
        "Step 2 — Join SCPAI",
        "         [Button: Join SCPAI -> your registration form]",
        "         Complete our club registration. Annual membership fee: [amount].",
        "",
        "Step 3 — Book and Play!",
        "         Book any session on OpenSports and show your member ID on arrival.",
        "         [Button: Book a Session -> OpenSports link below]",
        "",
        "Questions? Contact us -> /contact",
      ]),
      linkPara("OpenSports — Step 3 Book a Session button",OPENSPORTS_URL),
      sp(),

      h2("3.8  Join > Member Benefits"),
      copyBox("Benefits to highlight",[
        "Unlimited session access at all three venues",
        "PAA insurance cover included",
        "Member-only events and socials",
        "Access to club equipment during sessions",
        "Grading system — play with people at your level",
        "Club newsletter and updates",
        "Voting rights at the AGM",
        "Supporting a community not-for-profit",
      ]),
      pb(),

      h2("3.9  News  (NEW page)"),
      body("A blog/news page keeps the site feeling alive and helps with Google search rankings. In Squarespace, this is the Blog page type — rename it News."),
      sp(),h3("Suggested post categories"),
      bulletBold("Club News: ","New venues, committee updates, AGM notices, policy changes."),
      bulletBold("Session Wrap-ups: ","Short posts after a tournament or event with a couple of photos."),
      bulletBold("Player Spotlights: ","Interview-style posts featuring a member."),
      bulletBold("Tips & Skills: ","Short technique tips for beginners and intermediate players."),
      bulletBold("Visiting Players: ","Welcome notes for touring players or reports from away matches."),
      sp(),nb("Even 1-2 posts per month signals to Google and visitors that the club is active."),
      sp(),

      h2("3.10  Gallery  (NEW page — recommended)"),
      body("A photo gallery does more for community feel than almost any other page. Squarespace's Gallery page type handles layout automatically."),
      bullet("Ask members to send photos via the Contact page or at sessions."),
      bullet("Use Grid or Slideshow layout."),
      bullet("Caption each photo with venue name and approximate date."),
      bullet("Add a CTA at the bottom: Want to be in the next photo? Book a session."),
      linkPara("Book a session CTA on Gallery page",OPENSPORTS_URL),
      pb(),

      // S4
      h1("4.  SEO & Meta Descriptions"),
      sp(),
      body("Squarespace lets you set a custom meta description and browser-tab title for every page. Currently most pages have blank meta descriptions. Below are ready-to-paste values."),
      sp(),
      twoCol([
        ["Page","Suggested Meta Description (155 characters max)"],
        ["Homepage","Play pickleball on the Sapphire Coast NSW. 3 venues in Bermagui, Bega & Pambula. $10 per session, all equipment provided. Beginners welcome."],
        ["Our Story","Learn about the Sapphire Coast Pickleball Association — a volunteer-run, not-for-profit community club in NSW."],
        ["What is Pickleball?","New to pickleball? Discover how it's played, how it started, and why it's one of the world's fastest-growing sports."],
        ["Find a Session","Find and book a pickleball session near you. Venues in Bermagui, Bega & Pambula. All levels welcome, 18+, $10 per session."],
        ["Become a Member","Join the Sapphire Coast Pickleball Association. Unlimited sessions, PAA insurance, and a welcoming community."],
        ["FAQs","Everything you need to know about playing pickleball on the Sapphire Coast — booking, membership, rules, and more."],
        ["News","Latest news, tournament results, tips, and updates from the Sapphire Coast Pickleball Association."],
        ["Contact","Get in touch with the Sapphire Coast Pickleball Association. We'd love to hear from you."],
      ]),
      sp(),nb("In Squarespace: each page > gear icon > SEO tab > fill in SEO Description and Browser Tab title. The Browser Tab field controls what appears in Google results — not the page heading."),
      pb(),

      // S5
      h1("5.  Squarespace Implementation Tips"),
      sp(),h2("5.1  Recommended Squarespace Features"),
      bulletBold("Fluid Engine: ","Use for all new pages — full drag-and-drop column control without code."),
      bulletBold("Accordion blocks: ","For FAQs, Rules, and the merged Etiquette page."),
      bulletBold("Summary blocks: ","For venue cards on the homepage and the News feed preview."),
      bulletBold("Button blocks: ","Primary button = filled/solid; secondary = outlined."),
      bulletBold("Blog page type: ","Rename to News — Squarespace handles all layout automatically."),
      bulletBold("Gallery page type: ","Built-in — create from the Pages panel."),
      bulletBold("Nav button (optional): ","Add a Book Now button to the top-right corner of the nav bar under Design > Navigation Style — very visible, drives bookings."),
      sp(),

      h2("5.2  OpenSports — Where to Link It"),
      body("Your OpenSports events page should appear prominently in multiple places. Here is a full summary:"),
      twoCol([
        ["Location","Link text to use"],
        ["Homepage hero — primary CTA","Book a Session"],
        ["Homepage venue cards","Book Now  /  View Sessions"],
        ["Find a Session — top button","View & Book Sessions on OpenSports"],
        ["Each venue card","Book a spot"],
        ["Visiting the Sapphire Coast box","Book a Session"],
        ["Become a Member — Step 3","Book a Session on OpenSports"],
        ["Gallery page — bottom CTA","Book a Session"],
        ["Nav bar top-right (optional)","Book Now"],
      ]),
      sp(),
      linkPara("OpenSports — SCPAI events page (use everywhere above)",OPENSPORTS_URL),
      sp(),

      h2("5.3  Colour & Font Consistency"),
      twoCol([
        ["Primary (buttons, headings)","Teal — #006D77 (or your existing brand teal)"],
        ["Accent (highlights, CTAs)","Gold — #E8C547"],
        ["Background","White #FFFFFF or light grey #F8F8F8"],
        ["Body text","Dark grey #333333 (softer on screen than pure black)"],
        ["Links","Teal (same as primary)"],
      ]),
      sp(),

      h2("5.4  Mobile Check"),
      body("Over 60% of community sports site visitors arrive on mobile. After any changes, use Design > Preview Mobile to check:"),
      bullet("Hero text is readable — not too small or cropped."),
      bullet("Buttons are large enough to tap easily."),
      bullet("Navigation collapses into a hamburger menu cleanly."),
      bullet("Venue cards stack vertically, not shrink to tiny columns."),
      sp(),

      h2("5.5  Page Naming Fixes"),
      navTable([
        ["About 1 (browser tab)","What is Pickleball?","Fix meta title for this page"],
        ["safety-etiquette-1 (URL slug)","safety-etiquette-and-conduct","Cleaner URL for merged page"],
        ["register-join-us-1 (URL slug)","existing-member","Distinguish from registration page"],
        ["documents-1 (URL slug)","agm-documents","Descriptive URL"],
        ["Sapphire Coast Pickleball - Home","Sapphire Coast Pickleball","Remove - Home from site title"],
      ]),
      pb(),

      // S6
      h1("6.  Quick Wins — Do These First"),
      sp(),
      body("Tackle these first — each takes less than an hour and has high impact:"),
      sp(),
      twoCol([
        ["Priority","Action"],
        ["1 — Urgent","Fix page meta titles (especially About 1 -> What is Pickleball?)"],
        ["2 — Urgent","Add OpenSports booking link + two CTA buttons to homepage hero"],
        ["3 — High","Remove duplicate Insurance page from The Association nav"],
        ["4 — High","Replace Camel Rock hero with a pickleball action photo"],
        ["5 — High","Merge Safety & Etiquette + Code of Conduct into one page"],
        ["6 — Medium","Add meta descriptions to all key pages"],
        ["7 — Medium","Add Quick Stats bar to homepage"],
        ["8 — Medium","Add 2-3 member testimonial quotes to homepage"],
        ["9 — Medium","Create a News/Blog page (even with 1 post to start)"],
        ["10 — Lower","Create a Gallery page and invite members to contribute photos"],
      ]),
      sp(),
      linkPara("OpenSports — all SCPAI events",OPENSPORTS_URL),
      sp(),
      nb("The entire revamp can be done page-by-page while the site stays live. Start with the homepage and navigation, then work through each section at your own pace."),
      sp(),
      new Paragraph({children:[new TextRun({text:"Good luck with the revamp — the Sapphire Coast has a great club and it deserves a website that shows it off.",size:22,font:"Arial",color:DGREY,italics:true})],alignment:AlignmentType.CENTER,spacing:{before:240,after:240}}),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/sessions/confident-zealous-gauss/mnt/outputs/SCPAI_Website_Revamp_Brief.docx', buf);
  console.log('Done.');
});
