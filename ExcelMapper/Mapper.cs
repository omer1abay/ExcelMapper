namespace ExcelMapper
{
    /// <summary>
    /// Simply, all data and column information is taken as string from excel. 
    /// The entity is created to be the same as the column names. 
    /// The property corresponding to the relevant column is found using reflection. 
    /// The convertion is performed and the data is returned.
    /// </summary>
    /// <typeparam name="TDestination"></typeparam>
    public static class Mapper<TDestination>
    {
        public static TDestination Map(List<string> source, List<string> columns)
        {
            try
            {
                var destination = Activator.CreateInstance<TDestination>(); //create instance 
                var columnsAndSource = source.Zip(columns, (s, c) => new { Source = s, Column = c }); //this list created for iteration simultaneously
                foreach (var item in columnsAndSource)
                {
                    var destinationProperties = typeof(TDestination).GetProperty(item.Column); //get the related property in entity
                    var destinationType = typeof(TDestination).GetProperty(item.Column)?.PropertyType; //get the related propertys type
                    if (destinationProperties != null)
                    {

                        if (destinationType.IsGenericType && destinationType.GetGenericTypeDefinition().Equals(typeof(Nullable<>))) //if the property is nullable
                        {
                            if (item == null)
                            {
                                return destination;
                            }

                            destinationType = Nullable.GetUnderlyingType(destinationType); //get the underlying type of nullable property
                        }

                        var tempValue = Convert.ChangeType(item.Source, destinationType); //convertion string to specified type of property
                        destinationProperties.SetValue(destination, tempValue); //set value
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