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

    if (formData.passPhotoData && formData.passPhotoData.indexOf(",") !== -1) {
      try {
        img1File = photoFolder.createFile(dataURItoBlob(formData.passPhotoData, "Pass_" + formData.refNo));
        img1Url = img1File.getUrl();
      } catch(e) { console.error("Img1 Save Error: " + e); }
    }

    if (formData.fullPhotoData && formData.fullPhotoData.indexOf(",") !== -1) {
      try {
        img2File = photoFolder.createFile(dataURItoBlob(formData.fullPhotoData, "Full_" + formData.refNo));
        img2Url = img2File.getUrl();
      } catch(e) { console.error("Img2 Save Error: " + e); }
    }

    sheet.appendRow([
      new Date(), formData.refNo, formData.postApplied, formData.salary, formData.contract, formData.fullName,
      formData.nat, formData.rel, formData.dob, formData.age, formData.pob, formData.marital,
      formData.childCount, formData.weight, formData.height, formData.complexion, formData.edu,
      formData.grade, formData.passNo, formData.issueDate, formData.issuePlace, formData.expiryDate,
      formData.en_p, formData.en_f, formData.en_fl, formData.ar_p, formData.ar_f, formData.ar_fl,
      formData.all_countries, formData.all_periods,
      formData.skill_doc, formData.skill_sew, formData.skill_bab, formData.skill_cla,
      formData.skill_coo, formData.skill_was, formData.skill_iro, img1Url, img2Url
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
    "Timestamp", "Ref No", "Post Applied", "Monthly Salary", "Contract Period", "Full Name",
    "Nationality", "Religion", "Date of Birth", "Age", "Place of Birth", "Marital Status",
    "No of Children", "Weight", "Height", "Complexion", "Education", "Grade",
    "Passport No", "Issue Date", "Issue Place", "Expiry Date",
    "English Poor", "English Fair", "English Fluent", "Arabic Poor", "Arabic Fair", "Arabic Fluent",
    "Previous Countries", "Previous Periods",
    "Skill Driving", "Skill Sewing", "Skill Baby Sitting", "Skill Cleaning",
    "Skill Cooking", "Skill Washing", "Skill Ironing", "Passport Photo URL", "Full Photo URL"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
}

function generatePresentationSlide(formData, img1File, img2File) {
  const TEMPLATE_ID = "11XcA8P1Lk471PzrYbuFelrKbVXmVNY8Ol3qwbeei-Vc"; 
  const DESTINATION_FOLDER_ID = "1j1aV83OyWrLHvPTGNOgGuSv_SD_cV8NT"; 
  
  try {
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const destFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const copy = templateFile.makeCopy("Application_" + formData.refNo + "_" + formData.fullName, destFolder);
    const presentation = SlidesApp.openById(copy.getId());
    const slide = presentation.getSlides()[0];

    const fields = {
      "{{ref.no}}": formData.refNo,
      "{{pass.no}}": formData.postApplied,
      "{{mon.sal}}": formData.salary,
      "{{con.per}}": formData.contract,
      "{{nameif}}": formData.fullName,
      "{{nat}}": formData.nat,
      "{{reli}}": formData.rel,
      "{{dob}}": formData.dob,
      "{{age}}": String(formData.age),
      "{{pob}}": formData.pob,
      "{{cis}}": formData.marital,
      "{{noc}}": String(formData.childCount),
      "{{wei}}": formData.weight,
      "{{hei}}": formData.height,
      "{{com}}": formData.complexion,
      "{{edq}}": formData.edu,
      "{{grd}}": formData.grade,
      "{{passNo}}": formData.passNo,
      "{{doi}}": formData.issueDate,
      "{{pos}}": formData.issuePlace,
      "{{dox}}": formData.expiryDate,
      "{{all_countries}}": formData.all_countries,
      "{{all_periods}}": formData.all_periods
    };

    for (let key in fields) {
      slide.replaceAllText(key, fields[key] || "");
    }

    const tags = ["en_p", "en_f", "en_fl", "ar_p", "ar_f", "ar_fl", "doc", "sew", "bab", "cla", "coo", "was", "iro"];
    tags.forEach(function(tag) {
      let value = (tag.startsWith("en_") || tag.startsWith("ar_")) ? formData[tag] : formData["skill_" + tag];
      slide.replaceAllText("{{" + tag + "}}", value === "Checked" ? "✔" : "");
    });

    if (img1File) {
      replaceImageInSlide(slide, "{{image1}}", img1File.getBlob());
    }
    if (img2File) {
      replaceImageInSlide(slide, "{{image2}}", img2File.getBlob());
    }

    presentation.saveAndClose();
  } catch (err) {
    console.error("Slide Error: " + err.toString());
  }
}

function replaceImageInSlide(slide, tag, blob) {
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    let shape = shapes[i];
    try {
      if (shape.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        let text = shape.getText().asString().trim();
        if (text.includes(tag)) {
          const top = shape.getTop();
          const left = shape.getLeft();
          const width = shape.getWidth();
          const height = shape.getHeight();
          
          shape.remove(); 
          const image = slide.insertImage(blob);
          
          const imgWidth = image.getWidth();
          const imgHeight = image.getHeight();
          const ratio = Math.min(width / imgWidth, height / imgHeight);
          
          const newWidth = imgWidth * ratio;
          const newHeight = imgHeight * ratio;
          
          const newLeft = left + (width - newWidth) / 2;
          const newTop = top + (height - newHeight) / 2;
          
          image.setLeft(newLeft);
          image.setTop(newTop);
          image.setWidth(newWidth);
          image.setHeight(newHeight);
          
          break; 
        }
      }
    } catch (e) {}
  }
}

function dataURItoBlob(dataURI, fileName) {
  const parts = dataURI.split(',');
  const decoded = Utilities.base64Decode(parts[1]);
  const contentType = parts[0].split(':')[1].split(';')[0];
  return Utilities.newBlob(decoded, contentType, fileName);
}
