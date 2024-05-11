
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.AddressableAssets;

[System.Serializable]
public class DGTableData
{
    public int Id;
}

public class DGTable<T> where T : DGTableData
{
    private readonly Dictionary<int, T> copiedData = new();

    public T Get(int id)
    {
        if (copiedData.ContainsKey(id) == false) return null;
        return copiedData[id];
    }

    public void Load(string path)
    {
        string jsonName = typeof(T).ToString();
        jsonName = jsonName.Substring(0, jsonName.Length - 3); // R, O, W
        var handle = Addressables.LoadAssetAsync<TextAsset>($"{path}/{jsonName}.json");
        handle.WaitForCompletion();
        TextAsset text = handle.Result;
        if (text == null) throw new FileLoadException();
        var items = FromJson($"{{\"Items\":{text.text}}}");
        copiedData.Clear();
        for (int i = 0; i < items.Length; i++)
        {
            copiedData.Add(items[i].Id, items[i]);
        }
    }

    [System.Serializable]
    public class Wrapper
    {
        public T[] Items;
    }

    private T[] FromJson(string textData)
    {
        return JsonUtility.FromJson<Wrapper>(textData).Items;
    }
}
