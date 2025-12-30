function enviarCorreos() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  let estadoCol = headers.indexOf("Estado");


  if (estadoCol === -1) {
    sheet.getRange(1, headers.length + 1).setValue("Estado");
    estadoCol = headers.length;
  }

  for (let i = 1; i < data.length; i++) {

    const nombre = data[i][0];      
    const email = data[i][1];       
    const whatsapp = data[i][3];    
    const estado = data[i][estadoCol];

    if (estado !== "Correo enviado") {


      MailApp.sendEmail({
        to: "freddysilvatuesta@gmail.com",
        subject: "¡Nuevo Lead Registrado en MYPE X!",
        htmlBody: `
          <p><strong>Nombre:</strong> ${nombre}</p>
          <p><strong>Teléfono:</strong> ${whatsapp}</p>
        `
      });

      const htmlProspecto = `
        <div style="font-family: Arial; max-width: 600px;">
          <p>Hola ${nombre},</p>

          <p>Gracias por registrarte. Pronto nos pondremos en contacto contigo.</p>

          <div style="margin-top: 20px;">
            <a href="https://www.fstnegocios.com"
               style="background:#6A1B9A;color:#fff;padding:10px 15px;
               text-decoration:none;border-radius:5px;margin-right:10px;">
              Visitar Negocios
            </a>

            <a href="https://wa.me/51949638568"
               style="background:#25D366;color:#fff;padding:10px 15px;
               text-decoration:none;border-radius:5px;">
              Chatear por WhatsApp
            </a>
          </div>
        </div>
      `;

      MailApp.sendEmail({
        to: email,
        subject: "Gracias por contactarnos",
        htmlBody: htmlProspecto
      });

      sheet.getRange(i + 1, estadoCol + 1).setValue("Correo enviado");
    }
  }
}
