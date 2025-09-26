using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Models
{
    public class Client
    {
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

        public Client(string name, string address, string bank, string iban,
                       string vatNumber = "", string cameraDeComert = "", string termenulDePlata = "")
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Client name is required.", nameof(name));
            
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("Client address is required.", nameof(address));
            
            if (bank == null)
                throw new ArgumentNullException(nameof(bank), "Bank cannot be null.");
            
            if (iban == null)
                throw new ArgumentNullException(nameof(iban), "IBAN cannot be null.");
            
            _name = name;
            _address = address;
            _bank = bank;
            _iban = iban;
            _vatNumber = vatNumber ?? string.Empty;
            _cameraDeComert = cameraDeComert ?? string.Empty;
            _termenulDePlata = termenulDePlata ?? string.Empty;
        }

        public override string ToString()
        {
            return $"Client Information:\n" +
                   $"Name: {Name}\n" +
                   $"Address: {Address}\n" +
                   $"Bank: {Bank}\n" +
                   $"IBAN: {IBAN}\n" +
                   $"VAT: {VATNumber}\n" +
                   $"Camera de Comert: {CameraDeComert}\n" +
                   $"Termenul de plata: {TermenulDePlata}";
        }
    }
}