/**
 * Generate formatted LOI as PDF for a specific recipient
 * Returns shareable PDF URL.
 */
function generateLOI_(recipient) {
  // safe data + defaults
  var data = Object.assign({}, recipient, {
    Date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy"),
    ProjectName: "University of Santo Tomas Software Engineering Team",
    UniversityAddress1: "España Blvd, Sampaloc",
    UniversityAddress2: "Manila City, Metro Manila, 1000",
    ContactNumber: "0960 293 7255",
    YourName: "Shanley Clarence Lacanlale",
    Role: "Project Manager",
    Email: "shanleyclarence.lacanlale.cics@ust.edu.ph"
  });

  function safeName(name) {
    if (!name) return "client";
    return name.toString().replace(/[\\\/:*?"<>|]/g, "").slice(0, 120);
  }

  // Create Google Doc
  var doc = DocumentApp.create("LOI - " + safeName(data.CompanyName));
  var body = doc.getBody();

  // Clear the default empty paragraph instead of removing it
  if (body.getNumChildren() > 0) {
    body.getChild(0).asParagraph().clear();
  }

  // Set default paragraph style
  var normalAttrs = {};
  normalAttrs[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman';
  normalAttrs[DocumentApp.Attribute.FONT_SIZE] = 12;
  normalAttrs[DocumentApp.Attribute.BOLD] = false;
  body.setAttributes(normalAttrs);

  // helper to append paragraph with font and bold control
  function appendPara(text, bold, spacingAfterPoints) {
    var p = body.appendParagraph(text || "");
    var t = p.editAsText();
    t.setFontFamily('Times New Roman').setFontSize(12).setBold(!!bold);
    if (spacingAfterPoints !== undefined) p.setSpacingAfter(spacingAfterPoints);
    return p;
  }

  // helper to append a bullet list item
  function appendBullet(text) {
    var li = body.appendListItem(text || "");
    // ensure bullet glyph and default font
    li.setGlyphType(DocumentApp.GlyphType.BULLET);
    li.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);
    return li;
  }

  // === Team Members (names bold, roles normal) ===
  appendPara("Lacanlale, Shanley Clarence", true);
  appendPara("Project Manager", false);
  appendPara("Estrella, Andrea Helaena", true);
  appendPara("Business Analyst", false);
  appendPara("Cuevas, Aaron", true);
  appendPara("Systems Analyst", false);
  appendPara("Escuro, Nathaniel Joseph", true);
  appendPara("Backend Developer", false);
  appendPara("Caras, Mikyla Paula", true);
  appendPara("Frontend Developer", false);
  appendPara("Villegas, Alyssa Halle", true);
  appendPara("Frontend Developer", false);
  appendPara("Cruz, Joaquim Philippe", true);
  appendPara("Quality Assurance Officer", false);
  appendPara("Erasquin, Dan Angelo", true);
  appendPara("Quality Assurance Officer", false);

  appendPara("", false, 8); // small blank space

  // === Address block ===
  appendPara(data.ProjectName, false);
  appendPara(data.UniversityAddress1, false);
  appendPara(data.UniversityAddress2, false);
  appendPara(data.Date || "", false, 8);

  appendPara(data.CompanyName || "[Company Name]", false);
  appendPara(data.Address || "[Company Address]", false);
  appendPara("Manila City, Metro Manila", false, 8);

  // === Subject / Greeting ===
  var subj = appendPara("Re: Letter of Intent for Software Development Collaboration", true);
  subj.setSpacingAfter(6);

  appendPara("Dear " + (data.CompanyName || "[Company Name]") + ",", false, 8);

  // === Body paragraph(s) ===
  appendPara(
    "On behalf of the University of Santo Tomas Software Engineering Team, it is with great enthusiasm that I convey our intent to collaborate with " +
    (data.CompanyName || "[Company Name]") +
    " in crafting a software solution designed to serve your unique goals. As a dedicated team of Computer Science students from the University of Santo Tomas, we bring not only technical skills but also fresh perspectives, adaptability, and a drive to deliver meaningful outcomes.",
    false,
    8
  );

  // Proposed Areas (with bullets)
  appendPara("Proposed Areas of Collaboration:", true);
  appendBullet("Software Development: Our team is committed to building a solution that is thoughtfully designed, robust in performance, and intuitive for end-users. Beyond functionality, we aim to create a system that will streamline processes, optimize efficiency, and adapt to future needs.");
  appendBullet("Collaborative Engagement: We value open communication as the foundation of a successful project. To that end, we propose maintaining regular exchanges through meetings, progress reviews, and feedback sessions, ensuring transparency and alignment every step of the way.");
  appendPara("", false, 8);

  // Request for Partnership
  appendPara("Request for Partnership Confirmation", true);
  appendPara(
    "We respectfully seek your approval to formalize this collaboration. With your confirmation, we will prepare a comprehensive project outline detailing the objectives, deliverables, timelines, and shared responsibilities to guide our partnership.",
    false,
    8
  );

  // Closing
  appendPara("Closing Note", true);
  appendPara(
    "At the University of Santo Tomas Software Engineering Team, our mission is to go beyond delivering a product— we aspire to create solutions that make an impact. We believe this partnership holds the potential to bring real value to your organization while allowing us to apply our expertise in a meaningful and professional setting.",
    false,
    8
  );

  appendPara(
    "Should you have any questions or wish to discuss this further, please feel free to contact me at " +
    data.ContactNumber +
    " or via email at " +
    data.Email +
    ". We would be delighted to meet at your convenience.",
    false,
    8
  );

  appendPara(
    "Thank you very much for considering this opportunity. We look forward to working with you and contributing to your success.",
    false,
    12
  );

  // Signature
  appendPara("Yours sincerely,", false, 8);
  appendPara(data.YourName, true);
  appendPara(data.Role, false);
  appendPara(data.ProjectName, false);

  // Save & close
  doc.saveAndClose();

  // Export to PDF and store in "Generated LOIs" folder
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfBlob = docFile.getAs(MimeType.PDF);

  var folderIterator = DriveApp.getFoldersByName("Generated LOIs");
  var folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder("Generated LOIs");
  var pdfFile = folder.createFile(pdfBlob).setName("LOI - " + safeName(data.CompanyName) + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Trash the intermediate Google Doc to avoid clutter (optional)
  docFile.setTrashed(true);

  return pdfFile.getUrl();
}
