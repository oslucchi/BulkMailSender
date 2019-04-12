package BulkMailSender;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class SendEmailWithAttachment {
	private static String toNamecase(String s)
	{
		String sep = "";
		String retValue = "";
		s = s.toLowerCase();
		String[] items = s.split(" ");
		for(String item : items)
		{
			if (item.trim().length() == 0)
				continue;
			retValue = retValue + sep + item.substring(0, 1).toUpperCase().concat(item.substring(1));
			sep = " ";
		}
		return(retValue);
	}
	
	public static void main(String[] args) {
		int excelHeaderRow = -1; // la riga su cui si trova il column header per scegliere la colonna email 
		int startFrom = -1; // la prima riga da considerare su excel per valutare la spedizione
		int howMany = -1;
		int timeout = -1;
		int emailSendCheckbox = 0; // colonna in cui si trova il checkbox per inviare o meno la mail
		int emailSentCheckbox = 1; // colonna in cui si trova il checkbox quando una mail risulta inviata
		String emailSentBy = "Osvaldo.Lucchini@wedi.it";
		String excelEmailTOFieldName = "Mail";
		String excelEmailCCFieldName = "Mail-CC"; // "Agente;CapoArea"; // "Marco.Diterlizzi@wedi.it,Stefano.Broccoletti@wedi.it";
		String excelEmailBCCFieldName = "osvaldo.lucchini@wedi.it";
		String excelEmailAttachFilesFolder = ".\\docs\\condizioni2019\\Documenti\\";
		String excelEmailAttachFilesExt = ".pdf";
		String excelEmailAttachFieldName = "File";
//											".\\docs\\listinoExcel2019\\ListinoPubblicoIT.xlsx"; 
//											".\\docs\\marmomac\\ISO4211.pdf;" +
//											".\\docs\\marmomac\\Resistenza_alla_flessione.pdf";
		String signatureFilePath = "signature_osvaldo.png";
		String sheetName = "Condizioni";
		String excelFilePath = ".\\docs\\condizioni2019\\clienti.xls";
		String mailBodyPath = ".\\docs\\condizioni2019\\testomail.htm";
		String mailSubject = "Condizioni clienti 2019";
		
		final String mailServerUsername = "OLucchini"; //change accordingly
		final String mailServerPassword = "Qelppa12"; //change accordingly
		final String mailsServerHost = "mailserver.wedi.de";
		// final String mailsServerHost = "smtp.gmail.com";
		final int stopEvery = 50;
		final int stopFor = 5000;
		int countSent = 0;
		
		try {
			if (args.length < 5)
			{
				System.out.println("Need a record to start from in the excel file");
				System.out.println("Usage: SendEmailWithAttachment header_row record_to_start_from how_many timeout checkBoxIdx");
				System.exit(-1);
			}
			excelHeaderRow = Integer.parseInt(args[0]);
			startFrom = Integer.parseInt(args[1]);
			howMany = Integer.parseInt(args[2]);
			timeout = Integer.parseInt(args[3]);
			emailSendCheckbox = Integer.parseInt(args[4]);
		}
		catch(Exception e)
		{
			System.out.println("Exception " + e.getMessage() + " converting '" + args[0] + "' to int");
		}
		finally
		{
			if ((startFrom == -1) || (howMany == -1) || (timeout == -1))
			{
				System.out.println("Usage: SendEmailWithAttachment record_to_start_from how_many timeout");
				System.exit(-1);
			}
		}
		ReadFromExcel excel = null;
		try {
			excel = new ReadFromExcel(excelFilePath, excelHeaderRow,
									  sheetName, startFrom, excelEmailTOFieldName, excelEmailCCFieldName, 
									  excelEmailBCCFieldName, excelEmailAttachFieldName, excelEmailAttachFilesFolder, excelEmailAttachFilesExt);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		Properties props = new Properties();

		props.put("mail.smtp.host", mailsServerHost);
		props.put("mail.smtp.ssl.trust", mailsServerHost);
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");

		// wedi settings
		props.put("mail.smtp.port", "25");

		// Gmail settings
		//		props.put("mail.smtp.port", "587");

		props.put("mail.smtp.user", mailServerUsername);
		props.put("mail.smtp.password", mailServerPassword);

		BufferedReader br = null;
		StringBuilder mailBody = new StringBuilder();
		try {
			br = new BufferedReader(new FileReader(mailBodyPath));
			String line = br.readLine();

			while (line != null) {
				mailBody.append(line);
				mailBody.append("\n");
				line = br.readLine();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// Get the Session object.
		Session session = Session.getInstance(props,
				new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(mailServerUsername, mailServerPassword);
			}
		});

		int count = 0;
		try {
			while(excel.getNextRow() != null)
			{
				if ((excel.getField(emailSendCheckbox).compareTo("x") != 0) || 
					(excel.getEmail().compareTo("") == 0) ||
					((emailSentCheckbox >= 0) && (excel.getField(emailSentCheckbox).compareTo("*") == 0)))
					continue;
				if (count++ == howMany)
				{
					System.out.println("Sent " + howMany + " emails. Reached the row " + 
									   excel.rowIdx + " on the file. Quitting the job");
					break;
				}
				
				if (countSent++ == stopEvery)
				{
					try
					{
						Thread.sleep(stopFor);
					}
					catch(Exception e)
					{
						;
					}
					countSent = 1;
				}
				
				System.out.print("Row " + excel.rowIdx + " - client " + excel.getEmail());
				// Create a default MimeMessage object.
				Message message = new MimeMessage(session);

				// Set From: header field of the header.
				message.setFrom(new InternetAddress(emailSentBy));

				if ((excel.getEmail() == null) || (excel.getEmail().trim().compareTo("") == 0))
					continue;

				// Set To: header field of the header.
				message.setRecipients(Message.RecipientType.TO,
							InternetAddress.parse(excel.getEmail()));
				//			InternetAddress.parse("osvaldo.lucchini@gmail.com"));
				
				if (excel.getEmailCCValue() != null)
				{
					message.setRecipients(Message.RecipientType.CC,
							InternetAddress.parse(excel.getEmailCCValue()));
				//			InternetAddress.parse("osvaldo.lucchini@wedi.it"));
				}
				
				if (excel.getEmailBCCValue() != null)
				{
					message.setRecipients(Message.RecipientType.BCC,
								InternetAddress.parse(excel.getEmailBCCValue()));
				}

				// Set Subject: header field
				message.setSubject(mailSubject);

				// Create a multipar message and the message part
				Multipart multipart = new MimeMultipart();
				BodyPart messageBodyPart = new MimeBodyPart();
				DataSource source;
				try
				{
					// Now set the actual message
					String body = mailBody.toString();
					body = body.replaceFirst("\\$Codice\\$",excel.getField(2));
					body = body.replaceFirst("\\$Azienda\\$", toNamecase(excel.getField(3)));
					body = body.replaceFirst("\\$Indirizzo\\$", toNamecase(excel.getField(5)));
					body = body.replaceFirst("\\$Citta\\$", toNamecase(excel.getField(6)));
					messageBodyPart.setContent(body, "text/html");
					multipart.addBodyPart(messageBodyPart);

					// Part two is the picture
					MimeBodyPart imagePart = new MimeBodyPart();
					imagePart.attachFile(".\\docs\\wedi.png");
					imagePart.setContentID("<wediLogo>");
					imagePart.setDisposition(MimeBodyPart.INLINE);
					multipart.addBodyPart(imagePart);

					// signature needed?
					if (signatureFilePath != null)
					{
						imagePart = new MimeBodyPart();
						imagePart.attachFile(".\\docs\\" + signatureFilePath);
						imagePart.setContentID("<signature>");
						imagePart.setDisposition(MimeBodyPart.INLINE);
						multipart.addBodyPart(imagePart);
					}
					
					// attachments needed?
					int y = 0;
					String[] attachments = excel.getFileAttachValue().split(";");					
					while((y < attachments.length) && attachments[y].compareTo("") != 0)
					{
						messageBodyPart = new MimeBodyPart();
						source = new FileDataSource(attachments[y]);
						messageBodyPart.setDataHandler(new DataHandler(source));
						messageBodyPart.setFileName(attachments[y].substring(attachments[y].lastIndexOf("\\")));
						multipart.addBodyPart(messageBodyPart);
						y++;
					}

					// Send the complete message parts
					message.setContent(multipart);

					System.out.print(" sending to " + excel.getField(1) + " " +
							excel.getEmailCCValue() + " " + excel.getEmailBCCValue());
					Transport.send(message);
					System.out.println(" - sent successfully....");
					excel.setSentFlag("*");
					Thread.sleep(timeout);
				}
				catch(Exception e)
				{
					System.out.println("Mail to " + excel.getEmail() + " not sent. Exception " + e.getMessage());
					e.printStackTrace();
				}
			}
		} catch (MessagingException e) {
			throw new RuntimeException(e);
		}
		finally
		{
			try
			{
				excel.writeChanges();
				System.out.println("Changes saved");
			}
			catch(Exception e)
			{
				System.out.println("Changes not written on the source file. Exception " + e.getMessage());
			}
		}
	}
}
