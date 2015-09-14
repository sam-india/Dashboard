package src;

import java.io.File;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
  
public class SendEmail {  
 public static void main(String[] args) {  
  
 
	String From = "sajal.samaiya.06@gmail.com";
	String to = "sajal.samaiya.06@gmail.com";
	String cc = "sajal.samaiya.06@gmail.com";
	String body = "";
	String host = "192.168.1.10";
	
	try {
	String path="D:\\Report\\";
	String filename="MyTest.xls";
	File f = new File(path+filename);
	Multipart multipart = new MimeMultipart();
	BodyPart msgbodypart = new MimeBodyPart();
	msgbodypart.setContent(body,"text/html");
	multipart.addBodyPart(msgbodypart);
	MimeBodyPart attachmentpart = new MimeBodyPart();
	DataSource source = new FileDataSource(filename);
	attachmentpart.setDataHandler(new DataHandler(source));
	attachmentpart.setFileName(f.getName());
	multipart.addBodyPart(attachmentpart);
	
	//Get the session object  
	Properties props = System.getProperties();
    props.setProperty("mail.smtp.host", host);   
    Session session = Session.getDefaultInstance(props);  
   
  //compose message  
    
   MimeMessage message = new MimeMessage(session);  
   message.setFrom(new InternetAddress(From));//change accordingly  
   message.addRecipient(Message.RecipientType.TO,new InternetAddress(to));
   message.addRecipient(Message.RecipientType.CC, new InternetAddress(cc));
   message.setSubject("My First Email Test");  
   message.setContent(multipart);
     
   //send message  
   Transport.send(message);  
  
   System.out.println("message sent successfully");  
   
  } catch (MessagingException e) {throw new RuntimeException(e);}  
   
 }  
} 