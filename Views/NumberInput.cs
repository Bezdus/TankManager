using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace TankManager.Views
{
    /// <summary>
    /// Общее поведение числовых полей в диалогах: выделение текста при входе,
    /// запись значения перед сохранением, поиск полей с ошибкой ввода
    /// </summary>
    internal static class NumberInput
    {
        /// <summary>
        /// Записать в источник значение поля с фокусом: по Enter (IsDefault) фокус остаётся в поле,
        /// а привязка по потере фокуса или с задержкой ещё не сработала
        /// </summary>
        public static void CommitFocusedTextBox()
        {
            if (Keyboard.FocusedElement is TextBox focusedBox)
                focusedBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        }

        public static bool HasValidationErrors(DependencyObject parent)
        {
            if (Validation.GetHasError(parent))
                return true;

            int count = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < count; i++)
            {
                if (HasValidationErrors(VisualTreeHelper.GetChild(parent, i)))
                    return true;
            }

            return false;
        }

        public static void ShowValidationWarning(Window owner)
        {
            MessageBox.Show(owner, "Некоторые значения не распознаны как числа (выделены красным). Исправьте их.",
                owner.Title, MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        /// <summary>
        /// Выделение всего текста при входе в поле, чтобы новое число набиралось поверх старого
        /// </summary>
        public static void OnGotKeyboardFocus(object sender)
        {
            if (sender is TextBox textBox && !textBox.IsReadOnly)
                textBox.SelectAll();
        }

        public static void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Без этого щелчок мышью сразу сбрасывает выделение, поставленное в OnGotKeyboardFocus
            if (sender is TextBox textBox && !textBox.IsKeyboardFocusWithin)
            {
                e.Handled = true;
                textBox.Focus();
            }
        }
    }
}
