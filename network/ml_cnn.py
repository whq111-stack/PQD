# -*- coding: utf-8 -*-
"""
Created on Mon Jun 28 19:51:05 2021

@author: gdx
"""

import torch
from torch import nn
#from cnn_atten_rnn import CNN

# ML-CNN 基线网络：四级一维 CNN 后接全连接多标签分类器。
class CNN(nn.Module):
    """
    input_shape: batchsize * 1 * 640
    output_shape: batchsize * num_labels
    """
    def __init__(self):
        """构造与原 LGAN 特征提取器相同的四级 CNN。"""
        super(CNN, self).__init__()
        self.stage1 = nn.Sequential(
            nn.Conv1d(in_channels=1, out_channels=16, kernel_size=3, stride=1,padding=1),
            nn.BatchNorm1d(16),
            nn.Conv1d(16, 16, 3, 1, 1),
            nn.BatchNorm1d(16),
            nn.ReLU(True),
            nn.MaxPool1d(kernel_size=2, stride=2),

            nn.Conv1d(16, 32, 3, 1, 1),
            nn.BatchNorm1d(32),
            nn.Conv1d(32, 32, 3, 1, 1),
            nn.BatchNorm1d(32),
            nn.ReLU(True),
            nn.MaxPool1d(2, 2),

            nn.Conv1d(32, 64, 3, 1, 1),
            nn.BatchNorm1d(64),
            nn.Conv1d(64, 64, 3, 1, 1),
            nn.BatchNorm1d(64),
            nn.ReLU(True),
            nn.MaxPool1d(2, 2),

            nn.Conv1d(64, 128, 3, 1, 1),
            nn.BatchNorm1d(128),
            nn.Conv1d(128, 128, 3, 1, 1),
            nn.BatchNorm1d(128),
            nn.ReLU(True),
            nn.MaxPool1d(2, 2)
        )

    def forward(self, x):
        """输出未池化的一维卷积特征图。"""
        x=x.float()
        x = self.stage1(x)
        return x


class ML_CNN(nn.Module):
    """仅使用共享 CNN 特征和全连接层的多标签分类基线。"""

    """
    input_shape: batchsize * 1 * 640
    output_shape: batchsize * num_labels
    tips : include sigmoid
    """
    def __init__(self):
        """构造 ML-CNN 分类器。"""
        super(ML_CNN, self).__init__()
        self.cnn = CNN()
        self.gap = nn.AdaptiveMaxPool1d(1)

        self.classifier = nn.Sequential(
            nn.Linear(128, 128),
            nn.ReLU(True),
            nn.Dropout(0.25),
            nn.Linear(128, 8)
            )

    def forward(self, x):
        """输出 8 个扰动/正常状态的 sigmoid 概率。"""
        x = self.cnn(x)
        x = self.gap(x)
        x = x.view(-1,128)
        x = self.classifier(x)
        x = torch.sigmoid(x)
        return x


