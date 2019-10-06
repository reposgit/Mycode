using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace reservation
{
    class Program
    {
        static object locker = new object();
        static void Main(string[] args)
        {
            var storage = new storage();
            Parallel.For(1, 11, test);
            Console.ReadLine();
        }

        static void test(int userid)
        {
            lock (locker)//резервировать один товар может только первый успевший поток -> блокировка от других потоков
            {
                storage.breakpoint = false;//разблокировка для выполнения после предыдущего потока
                Console.WriteLine("ID потока: {0}", userid);
                for (int i = 1; i <= 1000; i++)
                {   
                    if(storage.breakpoint)
                    {
                        break;
                    }                
                    if (storage.state == "Open")//если соединение открыто прерываем цикл - чтобы другой поток в это время не мог сделать изменения
                    {
                        break;
                    }
                    storage.reserve("товар", 1, userid);
                }
            }
        }
    }       
}