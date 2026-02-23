function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No data received in post request");
    }
    
    const formData = JSON.parse(e.postData.contents);
    const result = processForm(formData);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error("Post Error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processForm(formData) {
  if (!formData) {
    return { "status": "error", "message": "Form data is missing" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Submissions");
    
    if (!sheet) {
      sheet = ss.insertSheet("Submissions");
      addHeadings(sheet);
    } else if (sheet.getLastRow() === 0) {
      addHeadings(sheet);
    }
    
    const folderName = "ZOD_Photos";
    const folders = DriveApp.getFoldersByName(folderName);
    const photoFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    let img1File = null;
    let img2File = null;
    let img1Url = "No Photo";
    let img2Url = "No Photo";

    if (formData.image1 && formData.image1.indexOf(",") !== -1) {
      try {
        img1File = photoFolder.createFile(dataURItoBlob(formData.image1, "Pass_" + formData.refNo));
        img1Url = img1File.getUrl();
      } catch(e) { console.error("Img1 Save Error: " + e); }
    }

    if (formData.image2 && formData.image2.indexOf(",") !== -1) {
      try {
        img2File = photoFolder.createFile(dataURItoBlob(formData.image2, "Full_" + formData.refNo));
        img2Url = img2File.getUrl();
      } catch(e) { console.error("Img2 Save Error: " + e); }
    }

    sheet.appendRow([
      new Date(), 
      formData.refNo, 
      formData.paf, 
      formData.monsal, 
      formData.sex, 
      formData.nameif,
      formData.nat, 
      formData.rel, 
      formData.dob, 
      formData.age, 
      formData.pob, 
      formData.cis, 
      formData.noc, 
      formData.weight, 
      formData.height, 
      formData.complexion, 
      formData.edu,
      formData.lang, 
      formData.passno, 
      formData.doi, 
      formData.poi, 
      formData.dox,
      formData.all_countries, 
      formData.all_periods,
      formData.bas, 
      formData.cfe, 
      formData.dec, 
      formData.cle, 
      formData.iro, 
      formData.sew, 
      formData.coo, 
      formData.was, 
      formData.nur, 
      img1Url, 
      img2Url
    ]);

    generatePresentationSlide(formData, img1File, img2File);

    return { "status": "success", "message": "Submitted Successfully!" };
  } catch (e) {
    console.error("Critical Process Error: " + e.toString());
    return { "status": "error", "message": e.toString() };
  }
}

function addHeadings(sheet) {
  const headers = [
    "Timestamp", "Ref No", "Post Applied", "Monthly Salary", "Sex", "Full Name",
    "Nationality", "Religion", "Date of Birth", "Age", "Place of Birth", "Civil Status",
    "No of Children", "Weight", "Height", "Complexion", "Education", "Knowledge of Language",
    "Passport No", "Issue Date", "Issue Place", "Expiry Date",
    "Previous Countries", "Previous Periods",
    "Baby Sitting", "Caring for Elderly", "Decoration", "Cleaning",
    "Ironing", "Sewing", "Cooking", "Washing", "Nursing", "Passport Photo URL", "Full Photo URL"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
}

function generatePresentationSlide(formData, img1File, img2File) {
  const TEMPLATE_ID = "1NwSJUtojos_48JpRX2LPEkqo4z0vySGrcZ0CJ7GvEt0"; 
  const DESTINATION_FOLDER_ID = "1j1aV83OyWrLHvPTGNOgGuSv_SD_cV8NT"; 
  
  try {
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const destFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const copy = templateFile.makeCopy("Application_" + formData.refNo + "_" + formData.nameif, destFolder);
    const presentation = SlidesApp.openById(copy.getId());
    const slide = presentation.getSlides()[0];

    const fields = {
      "{{ref.no}}": formData.refNo,
      "{{paf}}": formData.paf,
      "{{mon.sal}}": formData.monsal,
      "{{sex}}": formData.sex,
      "{{nameif}}": formData.nameif,
      "{{nat}}": formData.nat,
      "{{reli}}": formData.rel,
      "{{dob}}": formData.dob,
      "{{age}}": String(formData.age),
      "{{pob}}": formData.pob,
      "{{cis}}": formData.cis,
      "{{noc}}": String(formData.noc),
      "{{wei}}": formData.weight,
      "{{hei}}": formData.height,
      "{{com}}": formData.complexion,
      "{{edq}}": formData.edu,
      "{{lang}}": formData.lang,
      "{{passNo}}": formData.passno,
      "{{doi}}": formData.doi,
      "{{poi}}": formData.poi,
      "{{dox}}": formData.dox,
      "{{all_countries}}": formData.all_countries,
      "{{all_periods}}": formData.all_periods
    };

    for (let key in fields) {
      slide.replaceAllText(key, fields[key] || "");
    }

    const skills = ["bab", "cfe", "dec", "cle", "iro", "sew", "coo", "was", "nur"];
    skills.forEach(function(skill) {
      slide.replaceAllText("{{" + skill + "}}", formData[skill] === "Checked" ? "✔" : "");
    });

    if (img1File) replaceImageInSlide(slide, "{{image1}}", img1File.getBlob());
    if (img2File) replaceImageInSlide(slide, "{{image2}}", img2File.getBlob());

    presentation.saveAndClose();
  } catch (err) {
    console.error("Slide Error: " + err.toString());
  }
}

function replaceImageInSlide(slide, tag, blob) {
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    let shape = shapes[i];
    if (shape.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      if (shape.getText().asString().includes(tag)) {
        const top = shape.getTop();
        const left = shape.getLeft();
        const width = shape.getWidth();
        const height = shape.getHeight();
        shape.remove(); 
        const image = slide.insertImage(blob);
        const ratio = Math.min(width / image.getWidth(), height / image.getHeight());
        const newWidth = image.getWidth() * ratio;
        const newHeight = image.getHeight() * ratio;
        image.setLeft(left + (width - newWidth) / 2).setTop(top + (height - newHeight) / 2).setWidth(newWidth).setHeight(newHeight);
        break; 
      }
    }
  }
}

function dataURItoBlob(dataURI, fileName) {
  const parts = dataURI.split(',');
  const decoded = Utilities.base64Decode(parts[1]);
  const contentType = parts[0].split(':')[1].split(';')[0];
  return Utilities.newBlob(decoded, contentType, fileName);
}
