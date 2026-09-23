# -*- coding: utf-8 -*-
"""
Created on Mon Jun 28 20:07:58 2021

@author: gdx
"""

import torch
from torch.utils.data import Dataset
from scipy import io
import numpy as np

# 单标签/多标签数据集读取器。
# 原始 .mat 文件的最后 label_num 列是标签，其余列是一维波形采样值。
class Mc_dataread(Dataset):
    """读取单标签 PQD 数据，供 MC-CNN、DCNN 和 CNN-LSTM 基线使用。"""

    def __init__(self,file_path,data_name,label_num):
        self.x,self.y = self.data_read(file_path,data_name,label_num)
        self.len = len(self.x)
    
    def __len__(self):
        return self.len
    
    def __getitem__(self, item):
        return self.x[item], self.y[item]

    def data_read(self,file_path,data_name,label_num):
        """加载 .mat 文件，并按波形特征与标签切分数据。"""
        data_path = file_path+data_name +'.mat'
        features_struct = io.loadmat(data_path)
        train_data = features_struct['data']
        #train_data = train_data.astype(np.float32)  #change the data type       
        lable_start = train_data.shape[1] - label_num
        x=train_data[:,0:lable_start]
        y=train_data[:,lable_start:]
        return self.shape_transform(x,y)
    
    def shape_transform(self,x,y):
        """将 NumPy 数组转换为 PyTorch 的 batch × channel × length 格式。"""
        x=torch.from_numpy(x)
        y=torch.from_numpy(y.astype(np.int64))
        x=x.reshape((x.shape[0],1,-1)) #sample_num * 1 *sample_dim
        return x,y


class Ml_dataread(Dataset):
    """读取多标签 PQD 数据，供原始 LGAN 和 ML-CNN 使用。"""

    def __init__(self,file_path,data_name,label_num):
        self.x,self.y = self.data_read(file_path,data_name,label_num)
        self.len = len(self.x)
    
    def __len__(self):
        return self.len
    
    def __getitem__(self, item):
        return self.x[item], self.y[item]

    def data_read(self,file_path,data_name,label_num):
        """加载多标签 .mat 文件，并按最后 label_num 列读取标签。"""
        data_path = file_path+data_name +'.mat'
        features_struct = io.loadmat(data_path)
        train_data = features_struct['data']
        #train_data = train_data.astype(np.float32)  #change the data type       
        lable_start = train_data.shape[1] - label_num
        x=train_data[:,0:lable_start] 
        y=train_data[:,lable_start:]
        return self.shape_transform(x,y)
    
    def shape_transform(self,x,y):
        """将标签前补充正常状态标签，形成原 LGAN 使用的 8 维标签。"""
        x=torch.from_numpy(x)
        y=torch.from_numpy(y.astype(np.int64))
        x=x.reshape((x.shape[0],1,-1)) #sample_num * 1 *sample_dim
        ll=torch.ones((x.shape[0],1))
        y=torch.cat((ll,y),dim=1)
        #y_r=torch.nn.functional.one_hot(y, num_classes=48)
        return x,y
