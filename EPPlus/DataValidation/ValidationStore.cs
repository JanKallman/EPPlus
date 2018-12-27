using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Manages list of <see cref="IExcelDataValidation"/> objects partially ordered by their addresses.
    /// It uses binary search to find matching validation, thus reducing search complexity to O(log(N)) instead of O(N).
    /// </summary>
    internal class ValidationStore
    {
        private struct ValidationInfo
        {
            public IExcelDataValidation Validation { get; }
            public ExcelAddressBase Address { get; }

            public ValidationInfo(IExcelDataValidation validation) : this()
            {
                this.Validation = validation;
                this.Address = validation.Address;
            }

            public ValidationInfo(ExcelAddressBase address) : this()
            {
                this.Validation = null;
                this.Address = address;
            }
        }

        /// <summary>
        /// Manages list of <see cref="IExcelDataValidation"/> objects partially ordered by address projection.
        /// Projection is defined by axis X (columns) or Y (rows).
        /// </summary>
        private class PartiallyOrderedValidationList
        {
            private static readonly IExcelDataValidation[] Empty = new IExcelDataValidation[0];

            private readonly List<ValidationInfo> items = new List<ValidationInfo>();
            private readonly IComparer<ValidationInfo> comparer;

            public PartiallyOrderedValidationList(IComparer<ValidationInfo> comparer)
            {
                this.comparer = comparer;
            }

            public void Add(IExcelDataValidation validation)
            {
                var item = new ValidationInfo(validation);
                var index = this.items.BinarySearch(item, this.comparer);

                if (index >= 0)
                {
                    this.items.Insert(index, item);
                }
                else
                {
                    index = ~index;

                    if (index == this.items.Count)
                    {
                        this.items.Add(item);
                    }
                    else
                    {
                        this.items.Insert(index, item);
                    }
                }
            }

            public bool Remove(IExcelDataValidation validation)
            {
                return this.items.RemoveAll(x => x.Validation.Equals(validation)) > 0;
            }

            public IExcelDataValidation[] FindSame(ExcelAddressBase address)
            {
                var item = new ValidationInfo(address);
                var index = this.items.BinarySearch(item, this.comparer);

                if (index >= 0)
                {
                    int left = index;
                    while (left >= 0 && this.comparer.Compare(this.items[left], item) == 0)
                    {
                        left -= 1;
                    }

                    left += 1;

                    int right = index;
                    while (right < this.items.Count && this.comparer.Compare(this.items[right], item) == 0)
                    {
                        right += 1;
                    }

                    right -= 1;

                    int resultLength = right - left + 1;
                    var result = new IExcelDataValidation[resultLength];

                    for (int i = 0; i < resultLength; i++)
                    {
                        result[i] = this.items[left + i].Validation;
                    }

                    return result;
                }
                else
                {
                    return Empty;
                }
            }
        }

        /// <summary>
        /// Compares X axis (columns) projections of two ranges x and y.
        /// 
        /// Range x is considered to be less than y if end row of x is less than start column of y:
        /// XXX
        /// XXX YYY
        ///     YYY
        ///       
        /// Range x is considered to be greater than y if start column of x is greater than end column of y:
        /// YYY
        /// YYY XXX
        ///     XXX
        ///     
        /// Otherwise, ranges are considered equal. This means their projections have non-empty intersection area:
        /// XXX
        /// XXX
        ///   YYY
        ///   YYY
        /// </summary>
        private class ExcelDataValidationColumnComparer : IComparer<ValidationInfo>
        {
            public int Compare(ValidationInfo x, ValidationInfo y)
            {
                var addressX = x.Address;
                var addressY = y.Address;

                if (addressX.Start.Column + addressX.Columns - 1 < addressY.Start.Column)
                {
                    return 1;
                }

                if (addressY.Start.Column + addressY.Columns - 1 < addressX.Start.Column)
                {
                    return -1;
                }

                return 0;
            }
        }

        /// <summary>
        /// Compares Y axis (rows) projections of two ranges x and y.
        /// 
        /// Range x is considered to be less than y if end row of x is less than start row of y:
        /// XXX
        /// XXX
        ///   YYY
        ///   YYY
        /// 
        /// Range x is considered to be greater than y if start row of x is greater than end row of y:
        /// YYY
        /// YYY
        ///   XXX
        ///   XXX
        /// Otherwise, ranges are considered equal. This means their projections have non-empty intersection area:
        /// XXX
        /// XXX YYY
        ///     YYY
        /// </summary>
        private class ExcelAddressBaseRowComparer : IComparer<ValidationInfo>
        {
            public int Compare(ValidationInfo x, ValidationInfo y)
            {
                var addressX = x.Address;
                var addressY = y.Address;

                if (addressX.Start.Row + addressX.Rows - 1 < addressY.Start.Row)
                {
                    return 1;
                }

                if (addressY.Start.Row + addressY.Rows - 1 < addressX.Start.Row)
                {
                    return -1;
                }

                return 0;
            }
        }

        private readonly PartiallyOrderedValidationList projectionsX = new PartiallyOrderedValidationList(new ExcelDataValidationColumnComparer());
        private readonly PartiallyOrderedValidationList projectionsY = new PartiallyOrderedValidationList(new ExcelAddressBaseRowComparer());

        public void Add(IExcelDataValidation validation)
        {
            this.projectionsX.Add(validation);
            this.projectionsY.Add(validation);
        }

        public bool Remove(IExcelDataValidation validation)
        {
            return this.projectionsX.Remove(validation) && this.projectionsY.Remove(validation);
        }

        public IEnumerable<IExcelDataValidation> FindCollisions(ExcelAddressBase address, IExcelDataValidation validatingValidation)
        {
            var collidingValidationsX = this.projectionsX.FindSame(address);
            var collidingValidationsY = this.projectionsY.FindSame(address);

            var collidingValidations = collidingValidationsX
                .Intersect(collidingValidationsY)
                .Where(x => string.Equals(x.Address.WorkSheet, address.WorkSheet));

            if (validatingValidation != null)
            {
                collidingValidations = collidingValidations.Except(new[] { validatingValidation });
            }

            return collidingValidations;
        }
    }
}
