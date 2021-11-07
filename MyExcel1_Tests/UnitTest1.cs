using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using MyExcel1;

namespace MyExcel1_Tests
{
    [TestClass]
    public class UnitTest1
    {
        public static Cell temp = new Cell();
        public static Dictionary<string, Cell> dict = new Dictionary<string, Cell>() { { "A1", temp } };
        public Calculator calculator = new Calculator(dict);
        string dataLost = "";
        [TestMethod()]
        public void EvaluateTest1() // Тестування унарного оператора -
        {
            string expr = "--1";
            double expectedRes = 1;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvaluateTest2() // Тестування "зайвих" операторів і вільного місця
        {
            string expr = "+                  -------1";
            double expectedRes = -1;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvaluateTest3() // Тестування ділення з виводом цілого числа(без остачі, бо програма працює з цілими числами) 
        {
            string expr = "7/3";
            double expectedRes = 2;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvaluateTest4() // Тестування складного виразу
        {
            string expr = "-(mod(7,2)+3*6^2)";
            double expectedRes = -109;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }

        [TestMethod()]
        public void EvaluateTest5() // Тестування пріоритетів знаків
        {
            string expr = "2-2+2+2*2";
            double expectedRes = 6;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }


        [TestMethod()]
        public void EvaluateTest6() // Тестування унарного оператора у звязку з піднесенням у степінь
        {
            string expr = "-(2)^2";
            double expectedRes = -4;
            double res = calculator.Evaluate(expr, "A1", ref dataLost);
            Assert.AreEqual(expectedRes, res);
        }
    }
} 
