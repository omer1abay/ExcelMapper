namespace ExcelMapper
{
    public static class Mapper<TDestination>
    {
        public static TDestination Map(List<string> source, List<string> columns)
        {
            try
            {
                var destination = Activator.CreateInstance<TDestination>();
                var columnsAndSource = source.Zip(columns, (s, c) => new { Source = s, Column = c });
                foreach (var item in columnsAndSource)
                {
                    var destinationProperties = typeof(TDestination).GetProperty(item.Column);
                    var destinationType = typeof(TDestination).GetProperty(item.Column)?.PropertyType;
                    if (destinationProperties != null)
                    {

                        if (destinationType.IsGenericType && destinationType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                        {
                            if (item == null)
                            {
                                return destination;
                            }

                            destinationType = Nullable.GetUnderlyingType(destinationType);
                        }

                        var tempValue = Convert.ChangeType(item.Source, destinationType);
                        destinationProperties.SetValue(destination, tempValue);
                    }
                }

                return destination;
            }
            catch (Exception)
            {

                return Activator.CreateInstance<TDestination>();
            }
        }
    }
}