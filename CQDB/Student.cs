namespace CQDB
{
    using System;

    internal class Student
    {
        private string course;
        private string courseSpeciality;
        private string id;
        private string layer;
        private string name;
        private string origin;
        private string speciality;
        private string term;

        public Student()
        {
            this.id = string.Empty;
            this.name = string.Empty;
            this.course = string.Empty;
            this.layer = string.Empty;
            this.speciality = string.Empty;
            this.courseSpeciality = string.Empty;
        }

        public Student(string id, string name, string origin, string layer, string speciality, string grade)
        {
            this.id = string.Empty;
            this.name = string.Empty;
            this.course = string.Empty;
            this.layer = string.Empty;
            this.speciality = string.Empty;
            this.courseSpeciality = string.Empty;
            this.Id = id;
            this.Name = name;
            this.Course = origin;
            this.Layer = layer;
            this.Speciality = speciality;
            this.CourseSpeciality = grade;
        }

        public override string ToString()
        {
            return string.Format("{0} | {1} | {2} | {3} | {4} | {5}", new object[] { this.CourseSpeciality, this.Layer, this.Course, this.Speciality, this.Id, this.Name });
        }

        public string Course
        {
            get
            {
                return this.course;
            }
            set
            {
                this.course = value;
            }
        }

        public string CourseSpeciality
        {
            get
            {
                return this.courseSpeciality;
            }
            set
            {
                this.courseSpeciality = value;
            }
        }

        public string Id
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }

        public string Layer
        {
            get
            {
                return this.layer;
            }
            set
            {
                this.layer = value;
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }

        public string Origin
        {
            get
            {
                return this.origin;
            }
            set
            {
                this.origin = value;
            }
        }

        public string Speciality
        {
            get
            {
                return this.speciality;
            }
            set
            {
                this.speciality = value;
            }
        }

        public string Term
        {
            get
            {
                return this.term;
            }
            set
            {
                this.term = value;
            }
        }
    }
}

