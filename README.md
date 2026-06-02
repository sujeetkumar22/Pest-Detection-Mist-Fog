# Mist-Fog-Assisted Real-Time Pest Detection and Classification Framework Using Deep Transfer Learning for Smart Agriculture

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![PyTorch](https://img.shields.io/badge/PyTorch-2.0+-ee4c2c.svg)](https://pytorch.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

This repository contains the complete implementation and documentation for the project *"A Mist-Fog-Assisted Real-Time Pest Detection and Classification Framework Using Deep Transfer Learning for Smart Agriculture"*.

The project addresses the critical smart farming challenges of rural network latency, bandwidth constraints, and dataset class imbalances. It proposes a decentralized, three-tier hierarchical Mist-Fog-Cloud computing layout that performs real-time visual classification of crop pests at the field boundary, using a customized, class-balanced **MobileNetV3-Large** convolutional neural network.

---

## 📌 Project Overview & Architecture

Remote farmlands typically suffer from unstable wireless network coverage, making cloud-centric computer vision architectures impractical. This project resolves this bottleneck by localizing data preprocessing and inference at the edge of the farm boundary, utilizing three distinct tiers:

```
                  +----------------------------------------------+
                  |                 CLOUD TIER                   |
                  |     (Global Model Retraining & Macro Logs)   |
                  +----------------------------------------------+
                                         ▲
                                         | WAN (Periodic Metadata Only)
                                         ▼
                  +----------------------------------------------+
                  |                  FOG TIER                    |
                  |     (Local Workstation / Raspberry Pi 4)     |
                  |  Runs: Class-Balanced MobileNetV3-Large Core  |
                  +----------------------------------------------+
                                         ▲
                                         | LAN / LoRa / Wi-Fi (Cropped ROI)
                                         ▼
+---------------------------------------------------------------------------------+
|                                    MIST TIER                                    |
|   +-------------------+      +-------------------+      +-------------------+   |
|   |   ESP32-CAM #1    |      |   ESP32-CAM #2    |      |   ESP32-CAM #K    |   |
|   | (Pest Trap/Capture) |    | (Pest Trap/Capture) |    | (Pest Trap/Capture) |   |
|   +-------------------+      +-------------------+      +-------------------+   |
+---------------------------------------------------------------------------------+
```

1. **Mist Layer (Edge Capturing)**: Ultra-low-power trap-mounted nodes (such as ESP32-CAM) capture raw visual feeds, apply localized region-of-interest (ROI) cropping to strip away background noise, and transmit cropped image payloads over short-range channels (Wi-Fi/LoRa).
2. **Fog Layer (Localized Processing)**: On-site edge workstations (such as Raspberry Pi 4) aggregate streams, manage queuing, and execute the lightweight deep transfer learning model to perform real-time pest diagnosis.
3. **Cloud Layer (Global Analysis)**: Receives periodic, highly compressed metadata reports (pest distribution and coordinates) rather than raw images, avoiding network congestion while supporting macro-level migration tracking and offline model retraining.

---

## 🧠 Algorithmic Core & Customizations

### 1. Lightweight Model Backbone
The framework optimizes a compressed **MobileNetV3-Large** architecture, which achieves a high accuracy-to-resource ratio through:
* **Depthwise Separable Convolutions**: Decouples standard convolutions into spatial filtering (depthwise) and channel mixing (pointwise), reducing the mathematical FLOP burden by approximately 8 to 9 times.
* **Inverted Bottlenecks & Squeeze-and-Excitation (SE)**: Optimizes feature representation by projecting mappings into high-dimensional space before pooling, applying channel-wise attention weights.
* **Hard-Swish Activation**: Replaces standard swish calculations with a piece-wise linear approximation ($\text{h-swish}(x) = x \frac{\text{ReLU6}(x+3)}{6}$), allowing execution on edge CPUs using basic bitwise shifting and addition.

### 2. Mitigating Dataset Class Imbalances
Real-world pest capturing is highly skewed toward dominant insect species. To prevent majority class dominance, we integrate:
* **Dynamic Class-Balanced Weights**: Incorporates a volume scaling parameter $\beta = (V-1)/V$ to calculate the effective number of samples $E_c = (1 - \beta^{n_c}) / (1 - \beta)$ for each target class.
* **Weighted Random Sampler**: Integrates directly into the PyTorch dataloader, adjusting sampling probability inversely against class frequencies ($w_c \propto 1/E_c$) to distribute gradient updates evenly across all 9 target categories.
* **Class-Balanced Objective Function**: Minimized during backpropagation to prevent bias:
  $$\mathcal{L}_{CB}(y, \hat{y}) = - \sum_{c=1}^{C} w_c \cdot y_c \cdot \log(\hat{y}_c)$$

---

## 📊 Dataset Specifications & Categories

The network is trained and evaluated on a multi-class agricultural pest dataset containing **9 distinct categories** reflecting fine-grained visual challenges (high intra-class similarities and noisy field backgrounds):

1. **Aphids** (very small, high-density groupings)
2. **Armyworm** (caterpillar, torso markings)
3. **Beetle** (hard-shelled wing cases, visual overlaps)
4. **Bollworm** (larva phase, color variations)
5. **Grasshopper** (elongated hind limbs, green/brown camouflage)
6. **Mites** (arachnid class, microscopic details)
7. **Mosquito** (long slender appendages, wing scales)
8. **Sawfly** (wasp-like segmented torso layout)
9. **Stem Borer** (wood-boring caterpillars)

---

## 🛠️ Installation & Quick Start

### Prerequisites
* Python 3.8 or higher
* PyTorch (preferably CUDA-enabled for training)
* Windows/Linux/macOS

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/pest-detection-mist-fog.git
cd pest-detection-mist-fog
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```
*(If `requirements.txt` is not present, install core packages: `pip install torch torchvision torchaudio numpy matplotlib pillow pandas openpyxl python-docx python-pptx pywin32`)*

### 3. Run Training / Evaluation
The core notebook includes the entire training pipeline (data augmentation, AdamW optimizer, Cosine Annealing scheduler, early stopping, and evaluation metrics):
```bash
jupyter notebook mobilenetv3-large-pest-dataset.ipynb
```

---

## 📈 Experimental Results & Performance

Empirical evaluations on a **Raspberry Pi 4 Model B (4GB RAM)** target node confirm the framework's edge readiness:

| Metrics Parameter | Value | Details |
| :--- | :--- | :--- |
| **Best Validation Accuracy** | **97.33%** | Achieved across all 9 classes (no majority bias) |
| **Edge Frame Rate** | **42.5 FPS** | Exceeds the real-time threshold |
| **Inference Latency** | **23.5 ms** | Single-frame execution duration |
| **Model Size on Disk** | **16.4 MB** | Highly compressed MobileNetV3 weight parameters |
| **Active Memory Footprint** | **44.2 MB** | Runs stably within edge memory constraints |
| **CPU Utilization** | **38.4%** | Leaves CPU room for parallel networking tasks |

### Model Benchmark Comparisons
Benchmarked on edge target (Raspberry Pi 4) under identical training conditions:

| Convolutional Architecture | Best Validation Accuracy | Inference Latency (ms) | Storage Size (MB) | Peak Memory Draw (MB) |
| :--- | :---: | :---: | :---: | :---: |
| **MobileNetV3-Large (Ours)** | **95.78%** | **5.73** | **16.07** | **1249.68** |
| **ResNet18** | 95.56% | 2.06 | 42.65 | 2404.91 |
| **EfficientNet-B0** | 96.67% | 9.07 | 19.34 | 2253.25 |
| **EfficientNetV2-S** | 97.33% | 16.40 | 81.39 | 3593.29 |

---

## 📁 Repository Structure

```
├── mobilenetv3-large-pest-dataset.ipynb   # Main Jupyter notebook containing training and evaluation loops
├── notebook52d30dfc4f.ipynb               # Helper notebook for dataset profiling and preprocessing
├── README.md                              # This file
│
├── Main research/                         # Research assets, paper drafts, and plots
│   ├── Research Paper Draft.pdf           # Draft of the research paper
│   ├── mobilenet_class_distribution.png   # Class frequency imbalance plot (Figure 2)
│   ├── mobilenet_curves.png               # Training/validation curves (Figure 3)
│   ├── mobilenet_cm.png                   # Confusion matrix heatmap (Figure 4)
│   ├── mobilenet_metrics_heatmap.png      # Precision-Recall-F1 score heatmap (Figure 5)
│   ├── mobilenet_metrics_boxplot.png      # Evaluation metrics distribution (Figure 6)
│   ├── mobilenet_latency_profile.png      # Local edge inference latency histogram (Figure 7)
│   └── pest_predictions_grid.png          # Visual sample predictions (Figure 8)
│
├── Main report/                           # Final M.Sc. Project Report and Presentation files
│   ├── Sujeet_DUCS_Major_Project_Final_Math_Fixed.docx  # Word Report (with built-up equations & aligned layout)
│   ├── Sujeet_DUCS_Major_Project_Final_Math_Fixed.pdf   # PDF version of the Project Report
│   └── Sujeet_DUCS_Major_Project_Presentation.pptx      # PowerPoint Presentation (Mint-Teal Slate Theme)
│
└── Trained Models/                        # Pre-trained and fine-tuned weight checkpoint models
    └── final_mobilenet_pest.pth           # Weight file for the optimized MobileNetV3-Large core (17.0 MB)
```

---

## 📄 License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
