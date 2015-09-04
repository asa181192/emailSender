using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace pendingAprobalNotification
{
    class email
    {
        public static void sendEMailThroughOUTLOOK(string mail,string persona , string ticket,int dia)
        {
            try
            {
                SqlConnections temp = new SqlConnections(); // Create the Outlook application.
                Random rand = new Random();
               int unique = rand.Next(1,999999999);
                Outlook.Application oApp = new Outlook.Application();

                Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                System.Threading.Thread.Sleep(9000); // test


               Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
               oMsg.HTMLBody = "La persona : " + persona + " cuenta aun con una solicitud pendiente de aprobar<br/>"
                   + "Favor de ingresar a la siguiente liga http://192.168.26.160/pendingApproval/?id=" +ticket+"&unique="+unique+" para aprobar o rechazar el ticket " + ticket
                   +"<br/><b>Atencion : Una vez que actualize el ticket la liga dejara de funcionar es una liga unica. <b/>";
             //  string liga = "http://192.168.26.160/pendingApproval/?id="+ticket + "&unique=" + unique; 
              //string query = "INSERT INTO ligas (WORKORDERID,link) VALUES ("+ticket+","+unique+")";
               // temp.insertTemp(query);
               oMsg.Subject = "Pending Aproval Ticket: "+ticket;
               Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
               Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(mail);
               oRecip.Resolve();

                string query = "select count (*) FROM ligas WHERE WORKORDERID = "+ticket;
                int valor = temp.consultaDia(query);

                if (valor == 0)
                {
                    query = "INSERT INTO ligas (WORKORDERID,LINK,nombre_status) VALUES (" + ticket + "," + unique + "," + dia + ")";
                    temp.insertTemp(query);
                    oMsg.Send();
                    oRecip = null;
                    oRecips = null;
                    oMsg = null;
                    oApp = null;
                    Console.WriteLine("correo enviado");
                }//SI esl query regresa 0 el ticket no existe en la tabla temporal por lo tannto se crea un nuevo ticket 
                else
                {
                    query = "select nombre_status FROM ligas WHERE WORKORDERID = "+ticket;
                    valor = temp.consultaDia(query);

                    if (dia != valor)
                    {

                        query = "UPDATE ligas SET nombre_status = " + dia + " WHERE workorderid = " + ticket;
                        temp.insertTemp(query);
                        oMsg.Send();
                        oRecip = null;
                        oRecips = null;
                        oMsg = null;
                        oApp = null;
                        Console.WriteLine("correo enviado");

                    }
                    else
                    {
                        Console.WriteLine("El usuario ya recibio un correo el dia de hoy");

                    }
                }//si el query regresa diferente de 0 es porque el ticket ya existe en la tabla 
           
            }//end of try block
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }//end of catch
        }//end of Email Method
    }
}
