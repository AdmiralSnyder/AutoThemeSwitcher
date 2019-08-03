using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoThemeSwitcher
{
    public static class TaskUtils
    {
#pragma warning disable VSTHRD100 // "async void"-Methoden vermeiden
        public static async void FireAndForget(this Task task)
#pragma warning restore VSTHRD100 // "async void"-Methoden vermeiden
        {
            try
            {
                await task;
            }
            catch (Exception e)
            {
                // log errors
            }
        }
    }
}
