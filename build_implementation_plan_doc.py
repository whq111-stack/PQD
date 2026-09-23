from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from build_pqd_plan_doc import (
    add_body,
    add_bullet,
    add_heading,
    add_page_break,
    add_table,
    configure_document,
    set_run_font,
)


ROOT = Path(__file__).resolve().parent
OUT_PATH = ROOT / "docs" / "项目改进实施计划.docx"


def build_plan():
    doc = Document()
    configure_document(doc)

    title = doc.add_paragraph(style="Title")
    title.add_run("项目改进实施计划")
    subtitle = doc.add_paragraph(style="Subtitle")
    subtitle.add_run("基于物理条件扩散增强与暂态敏感双路径网络的复合电能质量扰动识别方法")
    intro = doc.add_paragraph()
    intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = intro.add_run("适用项目 LGAN PQD 复合扰动识别")
    set_run_font(r, size=10.5, color="555555")

    add_heading(doc, "一 计划目标", 1)
    add_body(doc, "本计划以现有 LGAN 工程为基线，逐步实现一套直接处理一维 PQD 波形的复合扰动识别方法。最终模型由物理条件扩散增强、暂态敏感双路径主干和组件级复合扰动解码三部分构成。计划不采用二维图像转换，不加入监督式早停，也不把 ModernTCN 或 Mamba 的网络名称本身作为创新点。")
    add_body(doc, "所有新增实验必须能够回答明确问题：扩散增强是否改善非理想工况，双路径主干是否改善暂态扰动，组件级解码是否改善未见复合组合泛化。原始 LGAN 代码和权重保留，作为可复现基线。")

    add_heading(doc, "二 当前程序入口", 1)
    add_table(doc, ["任务", "入口命令", "关键文件", "说明"], [
        ("原始 LGAN 训练", "python cnn_atten_rnn_train.py", "cnn_atten_rnn_train.py", "读取 train_ml_300_rand 和 train_ml_100_rand，保存 cnn_atten_rnn35.pth"),
        ("原始 LGAN 测试", "python model_test.py", "model_test.py", "默认测试 cnn_atten_rnn35.pth 和 test_ml_100_rand"),
        ("ML-CNN 基线", "python ml_cnn_train.py", "ml_cnn_train.py", "多标签 CNN 对比模型"),
        ("CNN-LSTM 基线", "python mc_cnn_train.py", "mc_cnn_train.py", "当前脚本默认训练 CNN_LSTM"),
        ("BPMLL 基线", "python bpmll_train.py", "bpmll_train.py", "Keras 多标签对比模型"),
        ("传统多标签基线", "python skmlmodel.py", "skmlmodel.py", "MLKNN 和 MLHARAM 测试入口"),
    ], widths=[3.2, 4.2, 4.5, 4.6], font_size=8.7)
    add_body(doc, "network/cnn_atten_rnn.py 是原始模型定义，不是独立入口；utils/dataread.py 是数据读取模块；utils/metrics.py 是指标模块。新模型将使用独立入口 train_tdsc.py 和 evaluate_tdsc.py，避免覆盖原始基线。")

    add_heading(doc, "三 文件改造总览", 1)
    add_table(doc, ["阶段", "新增文件", "修改文件", "保留文件"], [
        ("基线整理", "无", "为核心文件补充中文注释", "原始训练、测试、网络和权重"),
        ("数据生成", "data/pqd_generator.py、data/split_compositions.py", "utils/dataread.py", "原始 .mat 数据"),
        ("扩散增强", "network/conditional_diffusion.py、train_diffusion.py、sample_diffusion.py", "utils/physical_losses.py", "原始模型不变"),
        ("TDSC 主干", "network/tdsc.py、transient_path.py、modern_tcn_adapter.py、gated_fusion.py", "无，建议独立开发", "network/cnn_atten_rnn.py 作为基线"),
        ("组件解码", "network/component_decoder.py、task_heads.py", "utils/physical_losses.py", "原始 Decoupling 和 BRNN 作为对照"),
        ("训练测试", "train_tdsc.py、evaluate_tdsc.py、configs/pqd_config.py", "utils/metrics.py", "model_test.py 作为原始测试入口"),
    ], widths=[3.1, 6.2, 5.0, 3.0], font_size=8.3)

    add_page_break(doc)
    add_heading(doc, "四 第一阶段 基线确认和代码整理", 1)
    add_heading(doc, "4.1 目标", 2)
    add_body(doc, "在任何新模块开发前，确认原始 LGAN 可以在当前环境中完成训练和测试，并保留一份基线结果。基线结果至少包括 exact-match accuracy、Hamming loss、Macro-F1、Micro-F1、各扰动类别 Recall、参数量和推理时间。")
    add_heading(doc, "4.2 操作", 2)
    add_bullet(doc, "运行 python cnn_atten_rnn_train.py，确认使用 train_ml_300_rand 训练、train_ml_100_rand 验证并生成模型权重。")
    add_bullet(doc, "运行 python model_test.py，确认 test_ml_100_rand 可以完成预测。")
    add_bullet(doc, "记录 Python、PyTorch、CUDA、GPU、随机种子和模型文件版本。")
    add_bullet(doc, "保留原始 network/cnn_atten_rnn.py，不在基线阶段改变其网络结构。")
    add_bullet(doc, "扩展 utils/metrics.py 时保留原指标定义，新增指标不得改变原始结果。")
    add_heading(doc, "4.3 验收标准", 2)
    add_body(doc, "原始模型能够完整训练和测试；训练集、验证集和测试集路径明确；模型权重可加载；至少运行三个随机种子或重复实验；所有指标保存到结构化结果文件。")

    add_heading(doc, "五 第二阶段 数据集和物理条件扩展", 1)
    add_heading(doc, "5.1 数据接口", 2)
    add_body(doc, "现有 Ml_dataread 只返回波形和标签。新数据集类应返回字典或结构化对象，字段包括 waveform、component_labels、mechanism_labels、start_end、severity、snr、noise_type、domain_id 和 composition_id。原始 Ml_dataread 保留，新增 PQDDataset。")
    add_heading(doc, "5.2 数学数据生成", 2)
    add_table(doc, ["数据组", "内容", "用途"], [
        ("D0", "现有 640 点 .mat 数据", "原始 LGAN 同分布基线"),
        ("D1", "随机幅值、持续时间、频率、阻尼和发生位置", "未见参数外推"),
        ("D2", "对复合信号逐成分增加或删除", "成分重构和反事实一致性"),
        ("D3", "白噪声、彩色噪声、脉冲噪声、混合噪声", "噪声鲁棒性"),
        ("D4", "增益、偏置、量化、削顶、丢点、频率偏移", "跨设备和跨工况"),
        ("D5", "按组合而非按样本随机划分", "未见复合组合泛化"),
    ], widths=[2.4, 8.6, 5.8], font_size=8.8)
    add_heading(doc, "5.3 验收标准", 2)
    add_body(doc, "每个样本的类别、机制、参数和域信息可追溯；同源波形的不同噪声版本不会跨训练集和测试集泄漏；能够单独导出已见组合和未见组合测试集。")

    add_page_break(doc)
    add_heading(doc, "六 第三阶段 物理条件扩散增强", 1)
    add_heading(doc, "6.1 实现内容", 2)
    add_body(doc, "优先实现一维条件 DDPM 或 DDIM 采样模型。条件由扰动类别、物理机制、幅值、持续时间、起止位置、噪声类型和设备域组成。扩散模型只负责生成增强样本，不替代最终识别模型。")
    add_bullet(doc, "新增 network/conditional_diffusion.py，实现一维 U-Net、扩散步嵌入和条件嵌入。")
    add_bullet(doc, "新增 train_diffusion.py，单独训练和保存扩散模型。")
    add_bullet(doc, "新增 sample_diffusion.py，按指定组合、SNR 和设备域生成样本。")
    add_bullet(doc, "新增 utils/physical_losses.py，实现噪声预测、波形频域、幅值、频率和标签一致性约束。")
    add_bullet(doc, "使用规则检测器和冻结的基线分类器筛选生成样本，避免错误标签进入识别训练集。")
    add_heading(doc, "6.2 验收标准", 2)
    add_body(doc, "扩散样本满足设定的幅值、频率、持续时间和标签条件；生成样本具有比普通随机增强更丰富的域变化；加入扩散样本后，至少在低 SNR 或未见参数测试中带来可重复的下游收益。")

    add_heading(doc, "七 第四阶段 TDSC 主干网络", 1)
    add_heading(doc, "7.1 网络结构", 2)
    add_body(doc, "TDSC 直接接收 batch × 1 × 640 的一维波形，不转换为二维图像。局部暂态路径使用一阶差分、深度可分离卷积和空洞卷积；全局时序路径使用 ModernTCN；门控融合根据局部突变强度、估计 SNR 和频带统计量分配两条路径的权重。")
    add_table(doc, ["模块", "文件", "实现要点"], [
        ("局部暂态路径", "network/transient_path.py", "多尺度卷积、空洞率、差分残差、保留时间位置"),
        ("全局路径", "network/modern_tcn_adapter.py", "适配 ModernTCN 输入输出；Mamba 只作为备选"),
        ("门控融合", "network/gated_fusion.py", "根据信号统计量生成通道或时间门控"),
        ("总主干", "network/tdsc.py", "组织两条路径和融合输出"),
    ], widths=[3.7, 5.4, 7.7], font_size=8.9)
    add_heading(doc, "7.2 模型选择原则", 2)
    add_body(doc, "ModernTCN 作为首选全局骨干，原因是具有作者实现、适合一维时序并且工程复杂度较低。Mamba/SSM 作为可选对照，不作为项目阻塞项。InceptionTime、TCN、CNN-LSTM、DCNN、ML-CNN 和原始 LGAN 作为对比模型。网络名称不是论文创新，双路径、状态门控和物理任务耦合才是创新主体。")
    add_heading(doc, "7.3 验收标准", 2)
    add_body(doc, "TDSC 能在 D0 上独立训练；输入长度为 640 时输出尺寸稳定；局部路径、全局路径和融合路径均可单独消融；参数量和推理时间可统计；暂态扰动 Recall 相比原始 LGAN 有明确改善或至少不下降。")

    add_page_break(doc)
    add_heading(doc, "八 第五阶段 组件级解码和多任务识别", 1)
    add_heading(doc, "8.1 解码器", 2)
    add_body(doc, "第一版采用标签查询解码器。每类扰动拥有可学习查询向量，与 TDSC 特征执行交叉注意力，输出该成分的分类、定位、严重程度和重构特征。复合扰动按照无序集合处理，不依赖原始标签排列。扰动槽位解码器作为后续扩展。")
    add_bullet(doc, "新增 network/component_decoder.py，生成每类扰动组件特征。")
    add_bullet(doc, "新增 network/task_heads.py，提供分类、机制、定位、严重程度和未知风险头。")
    add_bullet(doc, "不继续使用原始 BRNN 的固定标签序列作为新模型的核心解码器。")
    add_heading(doc, "8.2 损失函数", 2)
    add_body(doc, "总损失由多标签分类、机制分类、起止位置、严重程度、成分重构、反事实一致性和域泛化损失组成。建议先训练分类和机制任务，再逐步加入重构和一致性任务，避免多任务梯度在初期互相干扰。")
    add_heading(doc, "8.3 验收标准", 2)
    add_body(doc, "模型能够输出具体扰动类别、机制、起止位置和严重程度；成分删除后的重构结果与对应反事实波形一致；未见复合组合测试中的 F1 高于普通标签分类器，并缩小已见组合和未见组合之间的性能差距。")

    add_heading(doc, "九 第六阶段 训练入口和评估入口", 1)
    add_table(doc, ["文件", "作用", "命令"], [
        ("train_tdsc.py", "分阶段训练 TDSC 和多任务识别模型", "python train_tdsc.py"),
        ("evaluate_tdsc.py", "执行所有新测试协议并保存结果", "python evaluate_tdsc.py"),
        ("configs/pqd_config.py", "统一管理路径、长度、类别、机制和超参数", "由训练和评估脚本导入"),
        ("utils/metrics.py", "增加 Macro-F1、Recall、IoU、AUROC 等指标", "保留原函数兼容性"),
    ], widths=[5.0, 8.0, 4.0], font_size=8.9)
    add_body(doc, "新入口不覆盖原始 cnn_atten_rnn_train.py 和 model_test.py。每次实验应保存配置、随机种子、权重路径、指标 JSON 和日志，保证论文图表可以重新生成。")

    add_page_break(doc)
    add_heading(doc, "十 实验计划", 1)
    add_table(doc, ["实验", "对比或测试内容", "核心指标", "通过标准"], [
        ("原始数据复现", "原始 LGAN、ML-CNN、CNN-LSTM、TDSC", "Accuracy、Hamming、Macro-F1", "新模型不低于原始 LGAN"),
        ("低 SNR 鲁棒性", "无噪声至 10 dB，白噪声和彩色噪声", "暂态 Recall、Macro-F1、性能下降", "暂态性能稳定提升"),
        ("未见组合", "训练未出现的双重和三重组合", "Seen/Unseen F1、Composition Gap", "Gap 明显缩小"),
        ("未见参数", "幅值、频率、阻尼、持续时间外推", "Macro-F1、参数 MAE", "证明物理参数外推"),
        ("跨设备域", "增益、偏置、量化、丢点和工频偏移", "跨域 F1、性能下降", "优于原始 LGAN"),
        ("定位与重构", "起止点、持续时间和组件重构", "MAE、IoU、频域误差", "组件输出具有物理解释"),
        ("未知拒识", "未知组合、未知噪声和未知波形", "AUROC、AUPR、FPR95、ECE", "降低强制误分类"),
    ], widths=[3.0, 6.0, 4.6, 3.8], font_size=8.3)

    add_heading(doc, "十一 消融实验", 1)
    add_table(doc, ["模型", "扩散增强", "TDSC", "组件解码", "重构一致性", "跨域未知"], [
        ("B0 原始 LGAN", "否", "否", "否", "否", "否"),
        ("B1", "否", "是", "否", "否", "否"),
        ("B2", "是", "是", "否", "否", "否"),
        ("B3", "是", "是", "是", "否", "否"),
        ("B4", "是", "是", "是", "是", "否"),
        ("B5 完整模型", "是", "是", "是", "是", "是"),
    ], widths=[3.8, 2.4, 2.3, 2.8, 3.3, 2.8], font_size=8.8)
    add_body(doc, "消融必须分别在原始同分布、20 dB、未见组合和跨域协议下报告，不能只报告随机测试集上的一个 accuracy。")

    add_heading(doc, "十二 风险和处理", 1)
    add_table(doc, ["风险", "处理方案"], [
        ("扩散样本标签不可靠", "使用数学参数、规则检测器和冻结分类器三重筛选；必要时降低扩散样本比例"),
        ("多任务训练不稳定", "分阶段训练、梯度裁剪、损失权重搜索和独立消融"),
        ("未见组合没有提升", "检查组合数据泄漏，增加成分增删反事实样本和组件重构约束"),
        ("ModernTCN 适配困难", "先用 TCN 或 InceptionTime 验证双路径逻辑，再替换全局骨干"),
        ("Mamba 依赖复杂", "只作为可选对照，不阻塞主模型"),
        ("缺少真实数据", "采用独立跨域仿真和少量半实物测试，并明确研究局限"),
    ], widths=[5.2, 12.1], font_size=8.8)

    add_heading(doc, "十三 最终交付物", 1)
    add_table(doc, ["交付物", "内容"], [
        ("代码", "数据生成器、扩散模型、TDSC、组件解码器、训练和评估入口"),
        ("数据", "D0 原始基线、D1 参数扩展、D2 反事实、D3 跨域和 D4 未见组合数据"),
        ("模型", "原始 LGAN 权重、TDSC 权重、完整模型权重和配置文件"),
        ("实验结果", "指标 JSON、训练日志、混淆矩阵、消融表、复杂度和可视化图"),
        ("论文材料", "网络结构图、数据生成流程图、损失函数、实验协议和创新点表述"),
    ], widths=[4.0, 13.3], font_size=9.0)
    add_body(doc, "完成本计划后，论文主线应保持为：一维原始波形输入，物理条件扩散扩展非理想数据，暂态敏感双路径主干提取局部和全局特征，组件级解码实现复合扰动识别与组合泛化。")

    doc.core_properties.title = "项目改进实施计划"
    doc.core_properties.subject = "LGAN PQD 复合扰动识别项目"
    doc.core_properties.author = ""
    doc.core_properties.comments = ""
    doc.save(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    build_plan()
