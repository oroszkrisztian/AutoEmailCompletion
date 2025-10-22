using System;

namespace EmailCompleteApp.Models
{
    public class Transportator
    {
        private int _id;
        private static int _nextId = 1;
        private string _name;
        private string _address;
        private string _bank;
        private string _iban;
        private string _vatNumber = string.Empty;
        private string _cameraDeComert = string.Empty;
        private string _termenulDePlata = string.Empty;

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

        public int Id
        {
            get => _id;
            set 
            {
               _id = value;
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

        public string Bank
        {
            get => _bank;
            set => _bank = value ?? string.Empty;
        }

        public string IBAN
        {
            get => _iban;
            set
            {
                if (!string.IsNullOrEmpty(value) && !System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Z0-9]+$"))
                    throw new ArgumentException("IBAN can only contain uppercase letters and numbers.", nameof(IBAN));
                
                _iban = value ?? string.Empty;
            }
        }

        public string VATNumber
        {
            get => _vatNumber;
            set => _vatNumber = value ?? string.Empty;
        }

        public string CameraDeComert
        {
            get => _cameraDeComert;
            set => _cameraDeComert = value ?? string.Empty;
        }

        public string TermenulDePlata
        {
            get => _termenulDePlata;
            set => _termenulDePlata = value ?? string.Empty;
        }

        public Transportator(string name, string address, string bank, string iban,
                           string vatNumber = "", string cameraDeComert = "", string termenulDePlata = "")
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Transportator name is required.", nameof(name));
            
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("Transportator address is required.", nameof(address));
            
            if (bank == null)
                throw new ArgumentNullException(nameof(bank), "Bank cannot be null.");
            
            if (iban == null)
                throw new ArgumentNullException(nameof(iban), "IBAN cannot be null.");
            
            _id = _nextId++;
            _name = name;
            _address = address;
            _bank = bank;
            _iban = iban;
            _vatNumber = vatNumber ?? string.Empty;
            _cameraDeComert = cameraDeComert ?? string.Empty;
            _termenulDePlata = termenulDePlata ?? string.Empty;
        }
    }
}