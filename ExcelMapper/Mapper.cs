namespace ExcelMapper
{
    public static class Mapper<TDestination>
    {
        public static TDestination Map(List<string> source, List<string> columns)
        {
            var destination = Activator.CreateInstance<TDestination>();
            var columnsAndSource = source.Zip(columns, (s, c) => new { Source = s, Column = c });
            foreach (var item in columnsAndSource)
            {
                var destinationProperties = typeof(TDestination).GetProperty(item.Column);
                var destinationType = typeof(TDestination).GetProperty(item.Column)!.PropertyType;
                if (destinationProperties != null)
                {
                    var tempValue = Convert.ChangeType(item.Source, destinationType);
                    destinationProperties.SetValue(destination, tempValue);
                }
            }

            return destination;
        }
    }
}
