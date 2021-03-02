namespace Spectrum.SpLoader.XMLSetting
{
    /// <summary>
    /// Файл мэппинга
    /// </summary>
    /// <remarks>Позволяет определить из каких столбцов файлов для сбора данных получать данные.</remarks>
    class Mapping
    {
        /// <summary>
        /// Имя мэппинга
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Указывает на столбец с указанием OmniClass
        /// </summary>
        public string Omni { get; set; }

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
        /// Примечание
        /// </summary>
        public string Note { get; set; }
    }
}
