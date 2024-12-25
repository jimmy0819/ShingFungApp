using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShinFung
{
    internal class SqlStruct
    {
       public static string BaseStruct= @"CREATE TABLE userdata (
                                        id INTEGER PRIMARY KEY,
                                        name STRING,
                                        zodiac STRING,
                                        age INTEGER,
                                        S_A INTEGER,
                                        S_B INTEGER,
                                        S_C INTEGER,
                                        S_C_type INTEGER,
                                        S_D INTEGER,
                                        S_E INTEGER,
                                        address STRING,
                                        D_A INTEGER,
                                        D_B INTEGER,
                                        phonenumber STRING,
                                        datetime INTEGER,
                                        sum INTEGER,
                                        target_name STRING,
                                        employee_name STRING,
                                        valid INTEGER
                                        );";
        //@"CREATE TABLE userdata (
        //                                id INTEGER PRIMARY KEY,0
        //                                name STRING,1
        //                                zodiac STRING,2
        //                                age INTEGER,3
        //                                S_A INTEGER,4
        //                                S_B INTEGER,5
        //                                S_C INTEGER,6
        //                                S_C_type INTEGER,7
        //                                S_D INTEGER,8
        //                                S_E INTEGER,9
        //                                address STRING,10
        //                                D_A INTEGER,11
        //                                D_B INTEGER,12
        //                                phonenumber STRING,13
        //                                datetime INTEGER,14
        //                                sum INTEGER,15
        //                                target_name STRING,16
        //                                employee_name STRING,17
        //                                valid INTEGER 18
        //                                );";
    }
}
