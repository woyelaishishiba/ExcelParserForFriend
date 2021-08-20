package com.helper.excel

fun merge(map1: Map<String, NodeInfo>, map2: Map<String, NodeInfo>): Map<String, NodeInfo> {
    val finalMap = HashMap<String, NodeInfo>()
    addAllEntries(map1, finalMap)
    addAllEntries(map2, finalMap)
    return finalMap
}

fun addAllEntries(map: Map<String, NodeInfo>, target: MutableMap<String, NodeInfo>) {
    for (entry in map) {
        target.putIfAbsent(entry.key, NodeInfo())
        entry.value.map.forEach { target[entry.key]?.addEntry(it.key, it.value) }
    }
}

class NodeInfo(val map: MutableMap<String, String> = HashMap()) {

    fun addEntry(key: String, value: String) {
        map[key] = value
    }
}
