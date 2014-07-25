using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;

namespace Priem
{
    public static class AsynchronousQueryExecutor
    {
        /// <summary>
        /// Выполняет отложенный LINQ-запрос.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query">Lazy Query (Convert into ToList() included)</param>
        /// <param name="callback">HandleResults Function</param>
        /// <param name="errorCallback">HandleError Function</param>
        public static void Call<T>(IEnumerable<T> query, Action<IEnumerable<T>> callback, Action<Exception> errorCallback)
        {
            Func<IEnumerable<T>, IEnumerable<T>> func =
                new Func<IEnumerable<T>, IEnumerable<T>>(InnerEnumerate<T>);
            IEnumerable<T> result = null;
            IAsyncResult ar = func.BeginInvoke(
                                query.ToList(),
                                new AsyncCallback(delegate(IAsyncResult arr)
                                {
                                    try
                                    {
                                        result = ((Func<IEnumerable<T>, IEnumerable<T>>)((AsyncResult)arr).AsyncDelegate).EndInvoke(arr);
                                    }
                                    catch (Exception ex)
                                    {
                                        if (errorCallback != null)
                                        {
                                            errorCallback(ex);
                                        }
                                        return;
                                    }
                                    //errors from inside here are the callbacks problem
                                    //I think it would be confusing to report them
                                    callback(result);
                                }),
                                null);
        }
        private static IEnumerable<T> InnerEnumerate<T>(IEnumerable<T> query)
        {
            foreach (var item in query) //the method hangs here while the query executes
            {
                yield return item;
            }
        }
    }
}
