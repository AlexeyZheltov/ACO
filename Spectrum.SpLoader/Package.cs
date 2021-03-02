using System.Collections.Generic;

namespace Spectrum.SpLoader
{
    /// <summary>
    /// Пакет для правильного добавления в HierarchyDictionary
    /// </summary>
    class Package
    {
        /// <summary>
        /// Сам эллемент данных
        /// </summary>
        public TargetItem Item { get; set; }

        /// <summary>
        /// Путь к расположению в иерархии
        /// </summary>
        public Stack<string> Path { get; set; }
    }
}
