using System.Collections.Generic;

namespace ACO.ExcelHelpers
{
    /// <summary>
    /// Класс иерархии данных
    /// </summary>
    /// <remarks>Создание подъуровней и расположение эллемента происходят автоматически</remarks>
    class HierarchyDictionary
    {
        /// <summary>
        /// Вложенные словари иерархии
        /// </summary>
        readonly Dictionary<string, HierarchyDictionary> _hierarchy = new Dictionary<string, HierarchyDictionary>();
      //  readonly List<TargetItem> _items = new List<TargetItem>();

        /// <summary>
        /// Сущность уровня иерархии
        /// </summary>
      //  public TargetItem Entity { get; set; }

        /// <summary>
        /// Добавляет эллемент в иерархию
        /// </summary>
        /// <remarks>Автоматически создавая нужные подъуровни</remarks>
        /// <param name="package">Пакет для добавления эллемента</param>
        //public void Add(Package package)
        //{
        //    if(package.Path.Count == 0)
        //    {
        //        package.Item.Level = 6;
        //        _items.Add(package.Item);
        //        return;
        //    }

        //    string temp = package.Path.Pop();
        //    int level = ++package.Item.Level;
        //    if(_hierarchy.TryGetValue(temp, out HierarchyDictionary hierarhyDictionary))
        //        hierarhyDictionary.Add(package);
        //    else
        //    {
        //        _hierarchy.Add(temp, new HierarchyDictionary
        //        {
        //            Entity = new TargetItem()
        //            {
        //                WorkName = temp,
        //                Unit = level == 5 ? "Комплект" : "Комплекс",
        //                Amount = "1",
        //                Level = level
        //            }
        //        });

        //        _hierarchy[temp].Add(package);
        //    }
        //}

        /// <summary>
        /// Получить последовательно все эллементы иерархии
        /// </summary>
        //public IEnumerable<TargetItem> Items()
        //{
        //    foreach(var hierarchy in _hierarchy.Values)
        //    {
        //        yield return hierarchy.Entity;

        //        foreach (var item in hierarchy.Items())
        //            yield return item;
        //    }

        //    foreach (var item in _items)
        //        yield return item;
        //}

        /// <summary>
        /// Производит нумерацию всех эллементов иерархии
        /// </summary>
        //public void Numeric(Numberer numberer, IProgressBarWithLogUI pb)
        //{
        //    foreach(var hItem in _hierarchy.Values)
        //    {
        //        if (pb.IsAborted) break;
        //        pb.SubBarTick();
        //        hItem.Entity.Number = numberer.GetNumber();
        //        hItem.Numeric(numberer.NextLevel(), pb);
        //    }

        //    foreach (var item in _items)
        //    {
        //        if (pb.IsAborted) break;
        //        pb.SubBarTick();
        //        item.Number = numberer.GetNumber();
        //    }
        //}

        //public int AllCount()
        //{
        //    int count = _hierarchy.Count;
        //    foreach (var hierarchy in _hierarchy.Values)
        //        count += hierarchy.AllCount();

        //    count += _items.Count;

        //    return count;
        //}
    }
}
