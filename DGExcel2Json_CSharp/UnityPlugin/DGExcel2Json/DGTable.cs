
using System;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.AddressableAssets;

[System.Serializable]
public class DGTableData
{
    public int Id;
}

public class DGTable<V> : Dictionary<int, V> where V : DGTableData
{
    public V Get(int id)
    {
        if (ContainsKey(id) == false) return null;
        return this[id];
    }

    public void Load(string path)
    {
        string jsonName = typeof(V).ToString();
        jsonName = jsonName.Substring(0, jsonName.Length - 3); // R, O, W
        var handle = Addressables.LoadAssetAsync<TextAsset>($"{path}/{jsonName}.json");
        handle.WaitForCompletion();
        TextAsset text = handle.Result;
        if (text == null) throw new FileLoadException();
        var items = FromJson($"{{\"Items\":{text.text}}}");
        this.Clear();
        for (int i = 0; i < items.Length; i++)
        {
            this.Add(items[i].Id, items[i]);
        }
    }

    [System.Serializable]
    public class Wrapper
    {
        public V[] Items;
    }

    private V[] FromJson(string textData)
    {
        return JsonUtility.FromJson<Wrapper>(textData).Items;
    }
}
