using System;

namespace EmailCompleteApp.Models
{
    public class Location
    {
        private int _id;
        private string _name;
        private string _address;
        private string _city;

        public int Id
        {
            get => _id;
            set => _id = value;
        }

        public string Name
        {
            get => _name;
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Name cannot be empty or whitespace.", nameof(Name));
                _name = value;
            }
        }

        public string Address
        {
            get => _address;
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Address cannot be empty or whitespace.", nameof(Address));
                _address = value;
            }
        }

        public string City
        {
            get => _city;
            set => _city = value;
        }

        public string DisplayAddress => $"{Address} ({City})";

        public Location(int id, string name, string address, string city)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Location name is required.", nameof(name));
            
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("Location address is required.", nameof(address));

            if (string.IsNullOrWhiteSpace(city))
                throw new ArgumentException("Location city is required.", nameof(city));

            _id = id;
            _name = name;
            _address = address;
            _city = city;
        }

        public override string ToString()
        {
            return DisplayAddress;
        }
    }
}
