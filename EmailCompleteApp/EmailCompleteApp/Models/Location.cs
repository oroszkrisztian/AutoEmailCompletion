using System;

namespace EmailCompleteApp.Models
{
    public class Location
    {
        private int _id;
        private static int _nextId = 1;
        private string _name;
        private string _address;

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

        public Location(string name, string address)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Location name is required.", nameof(name));
            
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("Location address is required.", nameof(address));
            
            _id = _nextId++;
            _name = name;
            _address = address;
        }

        public Location(int id, string name, string address)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Location name is required.", nameof(name));
            
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("Location address is required.", nameof(address));
            
            _id = id;
            _name = name;
            _address = address;
            
            if (id >= _nextId)
                _nextId = id + 1;
        }
    }
}
