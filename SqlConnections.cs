using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace pendingAprobalNotification
{
    class SqlConnections
    {

        private string connect;
        SqlConnection myConnection;
        SqlCommand cmd;

            public SqlConnections()
            {
          
                connect = System.Configuration.ConfigurationManager.ConnectionStrings["SqlConnection"].ConnectionString;
                myConnection = new SqlConnection(connect);
                cmd = new SqlCommand();
            }
            
            public void insertTemp (string query) {
             
                cmd.Connection = myConnection;
                cmd.CommandText = query;

                try
                {
                    myConnection.Open();
                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex ) {
                    Console.WriteLine(ex);

                }
                finally  {
                    myConnection.Close();
                }
                             
                   
                   
                

            }

            public int consultaDia(string query)
            {
                cmd.Connection = myConnection;
                cmd.CommandText = query;              
                object valor;

                try
                {
                    myConnection.Open();
                    valor = cmd.ExecuteScalar();
                    return int.Parse(valor.ToString());
                   
                   
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    return -1;
                }
                finally
                {
                    myConnection.Close();
                  
                }
            }

            public void consulta(string query)
            {
                cmd.Connection = myConnection;
                cmd.CommandText = query;
                SqlDataReader reader;
               

                try
                {   
                    myConnection.Open();
                    reader = cmd.ExecuteReader();
                    
                    while (reader.Read()) {
                           DateTime fecha = (DateTime)reader["Created Time"];

                      //  Console.WriteLine( (fecha-DateTime.Now).Days) ;
                      // Console.WriteLine(reader["correosubjefe"]+" "+reader["FIRST_NAME"] + " " + reader["Ticket"] + " " + reader["StatusID"] + " " + fecha.ToString("yyyy-MM-dd") + " Diferencia: " + (DateTime.Now - fecha).Days); 
                                          if ((DateTime.Now - fecha).Days==1)
                                       {
                                           if (reader["correosubjefe"].ToString()!="")
                                           {
                                              
                                               email.sendEMailThroughOUTLOOK(reader["correosubjefe"].ToString(), reader["FIRST_NAME"].ToString(), reader["TICKET"].ToString(),1);
                                           }
                                           
                                       }
                                       else if ((DateTime.Now - fecha).Days == 2)
                                       {
                                           if (reader["correojefe"].ToString()!="")
                                           {
                                               email.sendEMailThroughOUTLOOK(reader["correojefe"].ToString(), reader["FIRST_NAME"].ToString(), reader["TICKET"].ToString(),2);
                                           }
                                       }
                                       else if ((DateTime.Now - fecha).Days == 3)
                                       {
                                           if (reader["correodirector"].ToString()!="")
                                           {
                                               email.sendEMailThroughOUTLOOK(reader["correodirector"].ToString(), reader["FIRST_NAME"].ToString(), reader["TICKET"].ToString(),3);
                                           }
                                       }

                                      else if ((DateTime.Now - fecha).Days == 0)
                                       {

                                          if (reader["correousuario"].ToString() != "")
                                           {
                                               email.sendEMailThroughOUTLOOK(reader["correousuario"].ToString(), reader["FIRST_NAME"].ToString(), reader["TICKET"].ToString(),0);
                                           }
                                       }
                                       
                                         if ((DateTime.Now - fecha).Days == 4)
                                       {
                                           SqlConnections delete = new SqlConnections();
                                           query = "delete from ligas where workorderid = "+reader["TICKET"].ToString() ;
                                           delete.insertTemp(query);
                                           query = " update workorderStates set statusid = 3 where workorderid = " + reader["TICKET"].ToString();
                                           delete.insertTemp(query);
                                       }
                                         
                    }


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
                finally
                {
                    myConnection.Close();
                }
            }

    }
}
