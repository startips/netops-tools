#!/usr/bin/python
# -*- coding: utf-8 -*-
import re


def netSplit(net, input_maskBit=18, output_maskBit=27):  # 输入网段返回子网段List
    result = []  # 子网结果集
    network = net
    subnetBit = output_maskBit - input_maskBit
    subnetVar = 2 ** subnetBit
    networkList = network.split('.')
    networkBin = [int2bin(cell, 8) for cell in networkList]
    networkBin = ''.join(networkBin)[:input_maskBit]  # 转换成2进制网络号
    for maskBit in range(subnetVar):
        netMaskVar = networkBin + int2bin(maskBit, subnetBit)
        netMaskVar = netMaskVar.ljust(32, '0')  # 补充32位
        netMaskVarBin = re.findall(r'.{8}', netMaskVar)
        netMaskVarInt = [str(bin2int(cell)) for cell in netMaskVarBin]
        networkNew = '.'.join(netMaskVarInt)
        networkNew = f'{networkNew}/{output_maskBit}'
        result.append(networkNew)
    return result


def int2bin(value, bit):  # 十进制转换为二进制左侧不够位数补0 bit为位数
    binStr = bin(int(value))  # 二进制
    binStr = binStr.replace('0b', '')
    # binStr = binStr.zfill(3)  # 转换为三位二进制
    binEn = binStr.rjust(bit, '0')  # 二进制不足8位右边补0
    return binEn


def bin2int(value):  # 二进制转换为十进制
    intStr = int(value, 2)
    return intStr


def listReset(list):  # list重新排序 132465
    # result = []
    list_local = list
    for index in range(0, len(list_local), 3):
        index += 1
        try:
            list_local[index], list_local[index + 1] = list_local[index + 1], list_local[index]
        except IndexError:
            break
    return list_local


# 子网掩码长度转地址
def netmask_to_bit_length(netmask):
    return sum([bin(int(i)).count('1') for i in netmask.split('.')])


def bit_length_to_netmask(mask_int):
    bin_array = ["1"] * mask_int + ["0"] * (32 - mask_int)
    tmpmask = [''.join(bin_array[i * 8:i * 8 + 8]) for i in range(4)]
    tmpmask = [str(int(netmask, 2)) for netmask in tmpmask]
    return '.'.join(tmpmask)


if __name__ == '__main__':
    strs = 'xiaoxin122321'
    print(strs.removeprefix('xiaoxin'))
