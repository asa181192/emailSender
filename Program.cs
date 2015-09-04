using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace pendingAprobalNotification
{
    class Program
    {

        /*@Autor Alfredo Santiago Alvarado
         * 
         * 
         * 
        */

        static void Main(string[] args)
        {
               

           SqlConnections sql = new SqlConnections();
             string consulta = "Select empleados.correo AS CORREOUSUARIO,subjefes.correo AS CORREOSUBJEFE,jefes.CORREO AS CORREOJEFE,DIRECTORES.CORREO AS CORREODIRECTOR ,AAAUSER.FIRST_NAME,WORKORDER.WORKORDERID AS Ticket ,WORKORDERStates.StatusID,convert(date,dateadd(s,datediff(s,GETUTCDATE() ,getdate()) + (WORKORDER.CREATEDTIME/1000),'1970-01-01 00:00:00'),110) 'Created Time' " +
                                       "from WORKORDER  (nolock) " +
                                       "INNER JOIN  WORKORDERStates " +
                                       "ON WORKORDER.WORKORDERID = WORKORDERStates.WORKORDERID " +
                                       "Inner JOIN AAAUSER " +
                                       "ON WORKORDER.requesterid = AAAUSER.USER_ID " +
                                       "LEFT JOIN empleados " +
                                       "ON WORKORDER.requesterid = empleados.user_id " +
                                       "LEFT JOIN subjefes " +
                                       "ON empleados.subjefe_id = subjefes.user_id " +
                                       "LEFT JOIN jefes " +
                                       "ON empleados.jefe_id = JEFES.user_id " +
                                       "LEFT JOIN directores " +
                                       "ON empleados.director_id = directores.user_id " +
                                       "WHERE dateadd(s,datediff(s,GETUTCDATE() ,getdate()) + (WORKORDER.CREATEDTIME/1000),'1970-01-01 00:00:00') >= convert(varchar,CONVERT (date, GETDATE()-4),21) " +
                                       "AND WORKORDERStates.statusid = 301 ";
                sql.consulta(consulta);

             //  Console.WriteLine(sql.difDate("select DATEDIFF (day ,'2015-08-26' , CONVERT (date, SYSDATETIME())) AS diferencia"));
            
            
        //    email.sendEMailThroughOUTLOOK("lvazquez@transnetwork.com");
          
          
        } 
        
        
     
    }

   
}