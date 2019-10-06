using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace reservation
{
    class storage
    {
        public static string state;
        public static bool breakpoint;
        
        public static void reserve(string product, int amount, int userid)
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
                string sqlExpression = string.Format("SELECT count(*) FROM products where product='{0}'", product);                
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    state = connection.State.ToString();//проверка статуса соединения
                    try
                    {
                        SqlCommand cmd = new SqlCommand(sqlExpression, connection);
                        cmd.Connection = connection;
                        Int32 count = Convert.ToInt32(cmd.ExecuteScalar());
                        if (count == 0)//проверка есть такой товар в списке вообще или нет
                        {
                            Console.WriteLine("Такого товара нет в магазине");
                            connection.Close();
                            breakpoint = true;
                        }
                        else
                        {
                            sqlExpression = String.Format("SELECT amount from products where product='{0}'", product);//запрос количества товара на складе
                            cmd.CommandText = sqlExpression;
                            Int32 amountdb = Convert.ToInt32(cmd.ExecuteScalar());
                            if (amountdb < amount)
                            {
                                Console.WriteLine("Количество товара на складе меньше резервируемого, резервирование не выполнено");
                                connection.Close();
                                breakpoint = true;
                            }
                            else
                            {
                                amountdb = amountdb - amount;//расчет остатка
                                sqlExpression = String.Format("update [testbd].[dbo].[products] set [amount]={0}, [reservation]='process' where [product]='{1}'", amountdb, product);//обновление информации о товаре на складе (остаток и статус)
                                cmd.CommandText = sqlExpression;
                                cmd.ExecuteNonQuery();
                                sqlExpression = String.Format("update [testbd].[dbo].[reserv] set [amount]=[amount]+{0}, [reservation]='process' where [product]='{1}'", amount, product);//обновление информации о резервировании (сумма резерва и статус)
                                cmd.CommandText = sqlExpression;
                                cmd.ExecuteNonQuery();
                                breakpoint = false;
                                connection.Close();
                                Console.WriteLine("Пользователь: {0} /Зарезервировал товара '{1}': {2} шт /Остаток на складе: {3}",userid, product, amount, amountdb);                               
                            }
                        }
                    }                
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Исключение: {ex.Message}");
                    }
                    finally
                    {
                        state = connection.State.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Исключение: {ex.Message}");
            }
        }
    }
}