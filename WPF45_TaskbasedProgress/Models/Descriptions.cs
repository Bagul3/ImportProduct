using System;
using System.Text.RegularExpressions;

namespace ImportProducts.Models
{
    public class Descriptions
    {
        private readonly string _t2TRef;
        private readonly string _descriptio;
        private readonly string _description;
        private readonly string _bullet1;
        private readonly string _bullet2;
        private readonly string _bullet3;
        private readonly string _bullet4;
        private readonly string _bullet5;
        private readonly string _bullet6;
        private readonly string _bullet7;

        private Descriptions(DescriptionBuilder builder)
        {
            this._t2TRef = builder.T2TRef;
            this._descriptio = builder.Descriptio;
            this._description = builder.Description;
            this._bullet1 = builder.Bullet1;
            this._bullet2 = builder.Bullet2;
            this._bullet3 = builder.Bullet3;
            this._bullet4 = builder.Bullet4;
            this._bullet5 = builder.Bullet5;
            this._bullet6 = builder.Bullet6;
            this._bullet7 = builder.Bullet7;
        }

        public string GetT2TRef()
        {
            return _t2TRef;
        }

        public string GetDescriptio()
        {
            return _descriptio;
        }

        public string GetDescription()
        {
            return _description;
        }

        public string GetBullet1()
        {
            return _bullet1;
        }

        public string GetBullet2()
        {
            return _bullet2;
        }

        public string GetBullet3()
        {
            return _bullet3;
        }

        public string GetBullet4()
        {
            return _bullet4;
        }

        public string GetBullet5()
        {
            return _bullet5;
        }

        public string GetBullet6()
        {
            return _bullet6;
        }

        public string GetBullet7()
        {
            return _bullet7;
        }


        internal class DescriptionBuilder
        {
            internal string T2TRef;
            internal string Descriptio;
            internal string Description;
            internal string Bullet1;
            internal string Bullet2;
            internal string Bullet3;
            internal string Bullet4;
            internal string Bullet5;
            internal string Bullet6;
            internal string Bullet7;

            public DescriptionBuilder SetT2Tref(string reff)
            {
                T2TRef = Regex.Replace(reff, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetDescriptio(string descriptio)
            {
                Descriptio = Regex.Replace(descriptio, @"\t|\n|\r|[^0-9A-Za-z]+", ""); ;
                return this;
            }

            public DescriptionBuilder SetDescription(string description)
            {
                Description = Regex.Replace(description, @"\t|\n|\r|[^0-9A-Za-z]+", ""); ;
                return this;
            }

            public DescriptionBuilder SetBullet1(string bullet1)
            {
                Bullet1 = Regex.Replace(bullet1, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet2(string bullet2)
            {
                Bullet2 = Regex.Replace(bullet2, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet3(string bullet3)
            {
                Bullet3 = Regex.Replace(bullet3, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet4(string bullet4)
            {
                Bullet4 = Regex.Replace(bullet4, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet5(string bullet5)
            {
                Bullet5 = Regex.Replace(bullet5, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet6(string bullet6)
            {
                Bullet6 = Regex.Replace(bullet6, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public DescriptionBuilder SetBullet7(string bullet7)
            {
                Bullet7 = Regex.Replace(bullet7, @"\t|\n|\r|[^0-9A-Za-z]+", "");
                return this;
            }

            public Descriptions Build()
            {
                return new Descriptions(this);
            }

        }
    }
}
