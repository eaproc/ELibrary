using System;
using System.Linq.Expressions;

namespace ELibrary.Standard.Modules
{
    public static class basObjectNarrator
    {

        /// <summary>
        /// Usage: VB: Function() variableName, C# ()=> variableName
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetVariableFullName<T>(Expression<Func<T>> expr)
        {
            MemberExpression body = (MemberExpression)expr.Body;
            if (body.ToString().IndexOf("+") >= 0)
                return body.Member.Name; // REM it is a local variable
            return body.ToString().Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { ": " }, StringSplitOptions.RemoveEmptyEntries)[1];
        }

        /// <summary>
        /// Usage: VB: Function() variableName, C# ()=> variableName
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetVariableName<T>(Expression<Func<T>> expr)
        {
            MemberExpression body = (MemberExpression)expr.Body;
            return body.Member.Name;
        }

        /// <summary>
        /// Usage: Function() Me.MethodName(). if there is param, pass null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetMethodFullName<T>(Expression<Func<T>> expr)
        {
            MethodCallExpression body = (MethodCallExpression)expr.Body;
            return body.ToString().Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { ": " }, StringSplitOptions.RemoveEmptyEntries)[1];
        }

        /// <summary>
        /// Usage: Function() Me.MethodName(). if there is param, pass null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="expr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetMethodName<T>(Expression<Func<T>> expr)
        {
            MethodCallExpression body = (MethodCallExpression)expr.Body;
            return body.Method.Name;
        }

        public static string GetTypeFullName<T>(T obj)
        {
            return obj.GetType().FullName;
        }

        public static string GetTypeName<T>(T obj)
        {
            return obj.GetType().Name;
        }
    }
}