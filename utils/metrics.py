# -*- coding: utf-8 -*-
"""
Created on Mon Jun 28 20:13:06 2021

@author: gdx
"""

import torch
from sklearn import metrics
import numpy as np
from sklearn.metrics import accuracy_score, hamming_loss, \
    coverage_error,average_precision_score,label_ranking_loss 

# PyTorch 版本指标。
# 多标签输出通常是 sigmoid 概率，调用方应在需要离散预测时再进行阈值化。
class Mc_metrics():
    """单标签分类指标。"""

    def __init__(self):
        pass
    
    def acc(self,_input, target):
        target=target.squeeze()
        acc=float((_input.argmax(dim=1) == target).float().mean())
        return acc

#torch
class Ml_metrics():
    """原始 LGAN 使用的多标签指标集合。"""

    def __init__(self):
        pass
    
    def multi_acc(self, pred, true):
        """计算 exact-match accuracy：一条样本的所有标签都正确才算正确。"""
        pred=pred.clip(0,1).round()
        dpt = (pred - true).abs()
        acc = dpt.sum(dim=1).clip(0,1)
        return (1.-acc.mean()).item()
    
    def hamming_loss(self, pred, true):
        """计算所有样本和标签位置上的平均错误率。"""
        pred=pred.clip(0,1).round()
        dpt = (pred - true).abs()
        return dpt.mean().item()
    
    def rank_loss(self, pred, true):
        """计算正标签得分低于或等于负标签得分的比例。"""
        N = len(true)
        rankloss = 0.
        for i in range(N):
            index1 = torch.where(true[i] == 1)[0]
            index0 = torch.where(true[i] == 0)[0]
            tmp = 0
            for j in index1:
                for k in index0:
                    if pred[i,j] <= pred[i,k]:
                        tmp += 1
            rankloss += tmp * 1.0 / ((len(index1)) * len(index0))
        rankloss = rankloss / N
        return rankloss
    
    def one_errer(self, pred, true):
        """检查得分最高的标签是否属于真实标签集合。"""
        pred=pred.clip(0,1)
        pred = pred.argmax(dim=1)
        rr=true[torch.arange(0,pred.shape[0]),pred]
        return (1-rr.mean()).item()
    
    def coverage(self, pred, true):
        """计算覆盖误差，表示覆盖全部真实标签所需的排序范围。"""
        y_pred=pred.detach().numpy()
        y_true=true.numpy()
        c = metrics.coverage_error(y_true, y_pred)-1
        return c
    
    def average_precision(self, pred, true):
        """计算多标签平均精度。"""
        y_pred=pred.detach().numpy()
        y_true=true.numpy()
        ap = metrics.average_precision_score(y_true, y_pred) 
        return ap
    
# NumPy 版本指标，供 BPMLL、MLKNN 和 MLHARAM 基线使用。
class Ml_metrics_np():
    """基于 scikit-learn 的多标签指标封装。"""

    def __init__(self):
        pass
    
    def multi_acc(self, y_pred, y_true):
        y_pred = y_pred.round()
        acc = accuracy_score(y_true, y_pred)
        return acc
    
    def hamming_loss(self, y_pred, y_true):
        y_pred = y_pred.round()
        ham = hamming_loss(y_true, y_pred)
        return ham
    
    def rank_loss(self, y_pred, y_true):
        rl = label_ranking_loss(y_true, y_pred)
        return rl
    
    def one_errer(self, y_pred, y_true):
        pred = y_pred.argmax(axis=1)
        rr=y_true[np.arange(0,pred.shape[0]),pred]
        return 1-rr.mean()
    
    def coverage(self, y_pred, y_true):
        ce = coverage_error(y_true, y_pred)-1
        return ce
    
    def average_precision(self, y_pred, y_true):        
        ap = average_precision_score(y_true, y_pred)
        return ap
