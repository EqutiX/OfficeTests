using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadXlsx
{
    public class BaseObject
    {
        public int ID { get; set; }
    }

    public class ChildObject : BaseObject
    {
        private ChildObject()
        {

        }
        public ChildObject(string s)
        {

        }
    }
    public class TestClass<T>
    {
        public bool CreateNew<T>( T item ) where T : new()
        {
            return true;
        }
    }

    public class TestClass2
    {
        public bool CreateNew( BaseObject item )
        {
            return true;
        }
    }

    public class Foo
    {
        public void Bar()
        {
            var test1 = new TestClass();
            test1.CreateNew<ChildObject>(new ChildObject());

            var test2 = new TestClass2();
            test2.CreateNew(new ChildObject());
        }
    }
}
