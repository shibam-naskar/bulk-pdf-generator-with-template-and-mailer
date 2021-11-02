function myFunction() {
  var presentation = DriveApp.getFileById("GOOGLE SLIDE TEMPLATE ID");
  var folder = DriveApp.getFolderById("GOOGLE DRIVE FOLDER ID WHERE YOU WANT TO STORE SLIDES");
  var pdffolder = DriveApp.getFolderById("GOOGLE DRIVE FOLDER ID WHERE YOU WANT TO STORE GENERATED PDFS");
  var values = SpreadsheetApp.getActive().getDataRange().getValues();

  for(var i=1;i<values.length;i++){
    var copy = presentation.makeCopy(values[i][1],folder)
    var doc = SlidesApp.openById(copy.getId());
    var body = doc.getSlides()[0];
    body.replaceAllText("{{name}}",values[i][0]);
    // slide_array.push(copy.getId())
    doc.saveAndClose()
    Logger.log(`slide created id : ${copy.getId()}`)
    
    //this function is to store pdfs in foler
    
    generettePdf(copy.getId())
    
    //this function is to send email with pdf attachment
    sendEmail(index=i,id=copy.getId())
  }

  async function generettePdf(data){
    var blob = DriveApp.getFileById(data).getBlob();
    pdffolder.createFile(blob)
    Logger.log(`PDF created id: ${data}`)
  }

  async function sendEmail(index,id){
    var filepre = DriveApp.getFileById(id);
    var file = filepre.getAs(MimeType.PDF)
    
    //this is my html mailtemplate you can use your own template here
    
    var htmlBody = `
    <table border="0" cellpadding="0" cellspacing="0" class="nl-container" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #FFFFFF;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-1" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #fff;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 30px; padding-bottom: 20px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
      <table border="0" cellpadding="0" cellspacing="0" class="image_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="width:100%;padding-right:0px;padding-left:0px;">
      <div align="center" style="line-height:10px"><img alt="Image" src="https://pluralsight.imgix.net/course-images/securing-integrating-components-application-v1.png" style="display: block; height: auto; border: 0; width: 288px; max-width: 100%;" title="Image" width="288"/></div>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-2" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 50px; padding-bottom: 0px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
      <table border="0" cellpadding="10" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td>
      <div style="font-family: sans-serif">
      <div style="font-size: 14px; mso-line-height-alt: 16.8px; color: #555555; line-height: 1.2; font-family: Helvetica Neue, Helvetica, Arial, sans-serif;">
      <p style="margin: 0; font-size: 14px; text-align: center;"><span style="color:#03ff9f;"><strong><span style="font-size:42px;">Hellow ${values[index][0]}</span></strong></span></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td style="padding-bottom:30px;padding-left:10px;padding-right:10px;padding-top:10px;">
      <div style="font-family: sans-serif">
      <div style="font-family: Helvetica Neue, Helvetica, Arial, sans-serif; font-size: 12px; mso-line-height-alt: 14.399999999999999px; color: #1ac2f3; line-height: 1.2;">
      <p style="margin: 0; font-size: 14px; text-align: center;"><span style="font-size:58px;color:#0895f9;"><strong><span style="font-size:58px;">Congratulations you have successfully made it </span></strong></span></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="image_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="width:100%;padding-right:0px;padding-left:0px;">
      <div align="center" style="line-height:10px"><img alt="I'm an image" class="big" src="https://1.bp.blogspot.com/-y12_IgiqCXo/WjxiK4s367I/AAAAAAAAE8Q/jJTCZIDna18KqpGeu1ba93jGxcXqi4rqwCLcBGAs/s1600/image1.png" style="display: block; height: auto; border: 0; width: 640px; max-width: 100%;" title="I'm an image" width="640"/></div>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-3" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-image: url('https://htmlcolorcodes.com/assets/images/colors/baby-blue-color-solid-background-1920x1080.png'); background-position: top center; background-repeat: repeat; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 0px; padding-bottom: 0px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
      <table border="0" cellpadding="0" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td style="padding-bottom:10px;padding-left:10px;padding-right:10px;padding-top:45px;">
      <div style="font-family: sans-serif">
      <div style="font-size: 16px; font-family: Helvetica Neue, Helvetica, Arial, sans-serif; mso-line-height-alt: 19.2px; color: #555555; line-height: 1.2;">
      <p style="margin: 0; font-size: 16px; text-align: center;"><span style="color:#ffffff;font-size:16px;"><strong>HEY</strong></span></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td style="padding-bottom:10px;padding-left:30px;padding-right:30px;padding-top:10px;">
      <div style="font-family: sans-serif">
      <div style="font-size: 14px; font-family: Helvetica Neue, Helvetica, Arial, sans-serif; mso-line-height-alt: 21px; color: #555555; line-height: 1.5;">
      <p style="margin: 0; font-size: 14px; text-align: center; mso-line-height-alt: 24px;"><span style="font-size:16px;color:#ffffff;">You have successfully compleated ${values[index][6]==values[index][7]?'2 tracks':'1 track'} from 30 days of google cloud program .</span></p>
      <p style="margin: 0; font-size: 14px; text-align: center; mso-line-height-alt: 24px;"><span style="font-size:16px;color:#ffffff;">Its time to celibrate. And you can also show off your all skill badges in your social media . your all skill badges are listed here in the button bellow</span></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="button_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="padding-bottom:45px;padding-left:10px;padding-right:10px;padding-top:10px;text-align:center;">
      <div align="center">
      <!--[if mso]><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" style="height:56px;width:189px;v-text-anchor:middle;" arcsize="0%" strokeweight="1.5pt" strokecolor="#FFFFFF" fill="false"><w:anchorlock/><v:textbox inset="0px,0px,0px,0px"><center style="color:#ffffff; font-family:Arial, sans-serif; font-size:16px"><![endif]-->
      <a href=${values[index][5]}><div style="text-decoration:none;display:inline-block;color:#ffffff;background-color:transparent;border-radius:0px;width:auto;border-top:2px solid #FFFFFF;border-right:2px solid #FFFFFF;border-bottom:2px solid #FFFFFF;border-left:2px solid #FFFFFF;padding-top:10px;padding-bottom:10px;font-family:Helvetica Neue, Helvetica, Arial, sans-serif;text-align:center;mso-border-alt:none;word-break:keep-all;"><span style="padding-left:30px;padding-right:30px;font-size:16px;display:inline-block;letter-spacing:normal;"><span style="font-size: 16px; line-height: 2; word-break: break-word; mso-line-height-alt: 32px;"><strong>My Skill Badges</strong></span></span></div></a>
      <!--[if mso]></center></v:textbox></v:roundrect><![endif]-->
      </div>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-4" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 60px; padding-bottom: 15px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
      <table border="0" cellpadding="0" cellspacing="0" class="image_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="padding-bottom:45px;padding-top:10px;width:100%;padding-right:0px;padding-left:0px;">

      </td>
      </tr>
      </table>
      <table border="0" cellpadding="10" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td>
      <div style="font-family: sans-serif">
      <div style="font-family: Helvetica Neue, Helvetica, Arial, sans-serif; font-size: 12px; mso-line-height-alt: 14.399999999999999px; color: #555555; line-height: 1.2;">
      <p style="margin: 0; font-size: 14px; text-align: center;"><strong><span style="font-size:30px;color:#000000;">Tracks You Have Compleated</span></strong></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-5" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="50%">
      <table border="0" cellpadding="0" cellspacing="0" class="image_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="padding-right:20px;width:100%;padding-left:0px;padding-top:15px;">
      ${values[index][6]==6?`<div style="line-height:10px"><img alt="I'm an image" src="https://images.creativemarket.com/0.1.0/ps/6337907/1820/1214/m1/fpnw/wm0/10-.jpg?1556968407&amp;s=270c4896588c50827194df95343adade" style="display: block; height: auto; border: 0; width: 300px; max-width: 100%;" title="I'm an image" width="300"/></div>`:`<div></div>`}
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td style="padding-bottom:35px;padding-left:10px;padding-right:10px;padding-top:20px;">
      <div style="font-family: sans-serif">
      <div style="font-size: 16px; font-family: Helvetica Neue, Helvetica, Arial, sans-serif; mso-line-height-alt: 19.2px; color: #555555; line-height: 1.2;">
      ${values[index][6]==6?`<p style="margin: 0; font-size: 16px; text-align: center;"><span style="font-size:24px;color:#000000;"><strong>Cloud Engineering Track</strong></span></p>`:`<div></div`}
      </div>
      </div>
      </td>
      </tr>
      </table>
      </th>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="50%">
      <table border="0" cellpadding="0" cellspacing="0" class="image_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="padding-left:20px;width:100%;padding-right:0px;padding-top:15px;">
      ${values[index][7]==6?`<div align="right" style="line-height:10px"><img alt="I'm an image" src="https://images.creativemarket.com/0.1.0/ps/6337907/1820/1214/m1/fpnw/wm0/10-.jpg?1556968407&amp;s=270c4896588c50827194df95343adade" style="display: block; height: auto; border: 0; width: 300px; max-width: 100%;" title="I'm an image" width="300"/></div>`:`<div></div>`}
      </td>
      </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td style="padding-bottom:50px;padding-left:10px;padding-right:10px;padding-top:20px;">
      <div style="font-family: sans-serif">
      <div style="font-size: 16px; font-family: Helvetica Neue, Helvetica, Arial, sans-serif; mso-line-height-alt: 19.2px; color: #555555; line-height: 1.2;">
      ${values[index][7]==6?`<p style="margin: 0; font-size: 16px; text-align: center;"><span style="font-size:24px;color:#000000;"><strong>Data Science &amp; Machine Learning Track</strong></span></p>`:`<div></div`}
      </div>
      </div>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row row-6" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tbody>
      <tr>
      <td>
      <table align="center" border="0" cellpadding="0" cellspacing="0" class="row-content stack" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; color: #000000;" width="640">
      <tbody>
      <tr>
      <th class="column" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; vertical-align: top; padding-top: 5px; padding-bottom: 5px; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;" width="100%">
      <table border="0" cellpadding="0" cellspacing="0" class="icons_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="color:#9d9d9d;font-family:inherit;font-size:15px;padding-bottom:5px;padding-top:5px;text-align:center;">
      <table cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;" width="100%">
      <tr>
      <td style="text-align:center;">
      <!--[if vml]><table align="left" cellpadding="0" cellspacing="0" role="presentation" style="display:inline-block;padding-left:0px;padding-right:0px;mso-table-lspace: 0pt;mso-table-rspace: 0pt;"><![endif]-->
      <!--[if !vml]><!-->
      <table cellpadding="0" cellspacing="0" class="icons-inner" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block; margin-right: -4px; padding-left: 0px; padding-right: 0px;">

      </table>
      </table>
      <table border="0" cellpadding="10" cellspacing="0" class="text_block" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;" width="100%">
      <tr>
      <td>
      <div style="font-family: sans-serif">
      <div style="font-family: Helvetica Neue, Helvetica, Arial, sans-serif; font-size: 12px; mso-line-height-alt: 14.399999999999999px; color: #555555; line-height: 1.2;">
      <p style="margin: 0; font-size: 14px; text-align: center;"><strong><span style="font-size:30px;color:#000000;">Your certificate for Attendence is given bellow please download it</span></    strong></p>
      </div>
      </div>
      </td>
      </tr>
      </table>
      </td>
      </tr>
      </table>
      </td>
      </tr>
      </table>
      </th>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
      </td>
      </tr>
      </tbody>
      </table>
    `;
    
    
    MailApp.sendEmail({
      to: values[index][1],
      subject: "30 Days Of Gcp JGEC",
      htmlBody: htmlBody,
      attachments:file
    });
  }

  Logger.log("All PDF Created Successfully Script Made By SHIBAM NASKAR")
  
}
