namespace FormFiller.Helpers;

public static class ErrorHandler
{
    /// <summary>
    /// Executes an action and handles any exceptions that occur, allowing the program to continue.
    /// </summary>
    /// <param name="action">The action to execute.</param>
    /// <param name="onError">Optional: A callback to handle the exception.</param>
    public static void CatchError(Action action, Action<Exception>? onError = null)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            onError?.Invoke(ex);
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Executes a function and handles any exceptions that occur, allowing the program to continue.
    /// </summary>
    /// <typeparam name="T">The return type of the function.</typeparam>
    /// <param name="func">The function to execute.</param>
    /// <param name="onError">Optional: A callback to handle the exception.</param>
    /// <param name="defaultValue">The default value to return in case of an exception.</param>
    /// <returns>The result of the function or the default value if an exception occurs.</returns>
    public static T CatchError<T>(Func<T> func, Action<Exception>? onError = null, T defaultValue = default)
    {
        try
        {
            return func();
        }
        catch (Exception ex)
        {
            onError?.Invoke(ex);
            Console.WriteLine($"Error: {ex.Message}");
            return defaultValue;
        }
    }
}