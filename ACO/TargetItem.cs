namespace Spectrum.SpLoader
{
    /// <summary>
    /// Класс элемент данных.
    /// </summary>
    /// <remarks>Все переносимые из ресурсных файлов данные</remarks>
    class TargetItem
    {
        /// <summary>
        /// Omni
        /// </summary>
        public string OmniClass { get; set; }

        /// <summary>
        /// Наименование работ
        /// </summary>
        public string WorkName { get; set; }

        /// <summary>
        /// Маркировка/Обозначение
        /// </summary>
        public string Marking { get; set; }

        /// <summary>
        /// Материал
        /// </summary>
        public string Material { get; set; }

        /// <summary>
        /// Формат/Габаритные размеры/Диаметр (Ф) мм
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Тип, Марка, Обозначение документа, Опросного листа
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Артикул
        /// </summary>
        public string Article { get; set; }

        /// <summary>
        /// Производитель
        /// </summary>
        public string Maker { get; set; }

        /// <summary>
        /// Еденица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Кол-во
        /// </summary>
        public string Amount { get; set; }

        /// <summary>
        /// Комментарий
        /// </summary>
        public string Note { get; set; }

        /// <summary>
        /// Уровень маркера
        /// </summary>
        public int Level { get; set; }

        /// <summary>
        /// Номер
        /// </summary>
        public string Number { get; set; }
    }
}
