using System;

namespace TankManager.Core.Models
{
    public class Result<T>
    {
        public bool IsSuccess { get; }
        public T Value { get; }
        public string Error { get; }
        public Exception Exception { get; }

        private Result(bool isSuccess, T value, string error, Exception exception)
        {
            IsSuccess = isSuccess;
            Value = value;
            Error = error;
            Exception = exception;
        }

        public static Result<T> Success(T value) =>
            new Result<T>(true, value, null, null);

        public static Result<T> Failure(string error, Exception exception = null) =>
            new Result<T>(false, default, error, exception);
    }
}