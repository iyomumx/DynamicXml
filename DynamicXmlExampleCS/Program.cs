using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DynamicXml;

namespace DynamicXmlExampleCS
{
    class Program
    {
        static void Main(string[] args)
        {
            dynamic contact = new XDynamic("Contact");
            contact.Name = "Patrick Hines";
            contact.Phone = "206-555-0144";
            contact.Address = XDynamicExtensions.EmptyXDynamic;
            contact.Address.Street = "123 Main St";
            contact.Address.City = "Mercer Island";
            contact.Address.State = "WA";
            contact.Address.Postal = "68402";
            dynamic contact2 = new XDynamic("Contact");
            contact2.Name = "John Smith";
            contact2.Phone = "441-555-0206";
            contact2.Address = XDynamicExtensions.EmptyXDynamic;
            contact2.Address.Street = "777 West St";
            contact2.Address.City = "Mars";
            contact2.Address.State = "??";
            contact2.Address.Postal = "68412";
            dynamic contacts = new XDynamic("Contacts");
            contacts.Add((XElement)contact);
            contacts.Add((XElement)contact2);
            contacts.Contact[2] = XDynamicExtensions.EmptyXDynamic; //这里用超出1的Index来直接构建新的同名子级
            contacts.Contact[2].Name = "iyomumx";
            contacts.Contact[2].Phone = "111-555-1111";
            contacts.Contact[2].Address = XDynamicExtensions.EmptyXDynamic;
            contacts.Contact[2].Address.Street = "厦门大学漳州校区";
            contacts.Contact[2].Address.City = "漳州市（厦门市郊）";
            contacts.Contact[2].Address.State = "福建省";
            contacts.Contact[2].Address.Postal = "000000";
            //现在contacts里有三条记录
            var query = from c in (dynamic[])contacts.Contact               //不能用 as dynamic[]
                        where ((string)c.Address.Postal).Contains("0")      //强制转换为String和ToString()方法不同请注意
                        select c.Name;
            foreach (var name in query)
            {
                Console.WriteLine((string)name);
                name.Sub = "<Add A Sub Node>";
            }
            Console.ReadLine();
            Console.WriteLine(contacts.ToString());
            Console.ReadLine();
        }
    }
}
