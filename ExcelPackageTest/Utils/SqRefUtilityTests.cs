using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class SqRefUtilityTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void SqRefUtility_ToSqRefAddress_ShouldThrowIfAddressIsNullOrEmpty()
        {
            SqRefUtility.ToSqRefAddress(null);
        }

        [TestMethod]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommas()
        {
            // Arrange
            var address = "A1, A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.AreEqual("A1 A2:A3", result);
        }


        [TestMethod]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveCommasAndInsertSpaceIfNecesary()
        {
            // Arrange
            var address = "A1,A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.AreEqual("A1 A2:A3", result);
        }

        [TestMethod]
        public void SqRefUtility_ToSqRefAddress_ShouldRemoveMultipleSpaces()
        {
            // Arrange
            var address = "A1,        A2:A3";

            // Act
            var result = SqRefUtility.ToSqRefAddress(address);

            // Assert
            Assert.AreEqual("A1 A2:A3", result);
        }

        [TestMethod]
        public void SqRefUtility_FromSqRefAddress_ShouldReplaceSpaceWithComma()
        {
            // Arrange
            var address = "A1 A2";

            // Act
            var result = SqRefUtility.FromSqRefAddress(address);

            // Assert
            Assert.AreEqual("A1, A2", result);
        }
    }
}
