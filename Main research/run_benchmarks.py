import os
import time
import copy
import gc
import csv
import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader
from torchvision import models, transforms, datasets
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import accuracy_score, precision_recall_fscore_support, confusion_matrix, classification_report
import timm

# Hardware configuration
gc.collect()
torch.cuda.empty_cache()
os.environ["PYTORCH_ALLOC_CONF"] = "expandable_segments:True"
torch.backends.cudnn.benchmark = True 
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")

print(f"Using device: {device}")
if torch.cuda.is_available():
    print(f"Device Name: {torch.cuda.get_device_name(0)}")

# Paths
train_dir = r"E:\Sujeet\1.DU MSC\Semester 4\Sujeet Kumar DUCS Major Projest\Pest Dataset\archive\pest\train"
val_dir = r"E:\Sujeet\1.DU MSC\Semester 4\Sujeet Kumar DUCS Major Projest\Pest Dataset\archive\pest\test"
output_dir = r"E:\Sujeet\1.DU MSC\Semester 4\Main research"
models_dir = os.path.join(output_dir, "Trained Models")
os.makedirs(models_dir, exist_ok=True)

# Edge settings (hyperparameters from baseline)
img_size = 224
batch_size = 32
epochs = 25
patience = 5

class EarlyStopping:
    def __init__(self, patience=5):
        self.patience = patience
        self.counter = 0
        self.best_loss = None
        self.early_stop = False

    def __call__(self, val_loss):
        if self.best_loss is None:
            self.best_loss = val_loss
        elif val_loss > self.best_loss:
            self.counter += 1
            if self.counter >= self.patience:
                self.early_stop = True
        else:
            self.best_loss = val_loss
            self.counter = 0

# Transformations
train_tf = transforms.Compose([
    transforms.Resize((img_size, img_size)),
    transforms.RandomHorizontalFlip(),
    transforms.RandomRotation(20),
    transforms.ColorJitter(0.1, 0.1),
    transforms.ToTensor(),
    transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])
])

val_tf = transforms.Compose([
    transforms.Resize((img_size, img_size)),
    transforms.ToTensor(),
    transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])
])

# Datasets & Dataloaders
train_data = datasets.ImageFolder(train_dir, train_tf)
val_data = datasets.ImageFolder(val_dir, val_tf)
classes = train_data.classes
num_classes = len(classes)

train_loader = DataLoader(train_data, batch_size=batch_size, shuffle=True, num_workers=0, pin_memory=True)
val_loader = DataLoader(val_data, batch_size=batch_size, shuffle=False, num_workers=0, pin_memory=True)

print(f"Dataset summary:")
print(f"  Classes: {classes}")
print(f"  Train samples: {len(train_data)}")
print(f"  Test/Val samples: {len(val_data)}")

# Helper to load models
def get_model(model_name):
    if model_name == "MobileNetV3-Large":
        model = models.mobilenet_v3_large(weights='DEFAULT')
        in_features = model.classifier[-1].in_features
        model.classifier[-1] = nn.Linear(in_features, num_classes)
    elif model_name == "ShuffleNetV2 (0.5x)":
        model = models.shufflenet_v2_x0_5(weights='DEFAULT')
        in_features = model.fc.in_features
        model.fc = nn.Linear(in_features, num_classes)
    elif model_name == "GhostNet (1.0x)":
        # Load from timm
        model = timm.create_model('ghostnet_100', pretrained=True, num_classes=num_classes)
    elif model_name == "SqueezeNet":
        model = models.squeezenet1_1(weights='DEFAULT')
        model.classifier[1] = nn.Conv2d(512, num_classes, kernel_size=(1,1), stride=(1,1))
        model.num_classes = num_classes
    elif model_name == "ResNet18":
        model = models.resnet18(weights='DEFAULT')
        in_features = model.fc.in_features
        model.fc = nn.Linear(in_features, num_classes)
    elif model_name == "EfficientNet-B0":
        model = models.efficientnet_b0(weights='DEFAULT')
        in_features = model.classifier[1].in_features
        model.classifier[1] = nn.Linear(in_features, num_classes)
    elif model_name == "EfficientNetV2-S":
        model = models.efficientnet_v2_s(weights='DEFAULT')
        in_features = model.classifier[1].in_features
        model.classifier[1] = nn.Linear(in_features, num_classes)
    else:
        raise ValueError(f"Unknown model: {model_name}")
    return model.to(device)

# List of models to benchmark
model_names = [
    "MobileNetV3-Large",
    "ShuffleNetV2 (0.5x)",
    "GhostNet (1.0x)",
    "SqueezeNet",
    "ResNet18",
    "EfficientNet-B0",
    "EfficientNetV2-S"
]

results = []

for name in model_names:
    print("\n" + "="*50)
    print(f"Training and Evaluating Model: {name}")
    print("="*50)
    
    model = get_model(name)
    
    # Optimizer and Criterion (consistent with baseline)
    criterion = nn.CrossEntropyLoss()
    optimizer = optim.AdamW(model.parameters(), lr=1e-3, weight_decay=1e-4)
    scheduler = optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode='min', factor=0.5, patience=2)
    scaler = torch.amp.GradScaler('cuda')
    stopper = EarlyStopping(patience=patience)
    
    history = {'train_acc': [], 'val_acc': [], 'train_loss': [], 'val_loss': []}
    best_acc = 0.0
    best_model_wts = copy.deepcopy(model.state_dict())
    
    # Train Loop
    for epoch in range(epochs):
        start_t = time.time()
        
        # training phase
        model.train()
        t_loss, t_correct, t_total = 0.0, 0, 0
        for imgs, lbls in train_loader:
            imgs, lbls = imgs.to(device), lbls.to(device)
            optimizer.zero_grad()
            with torch.amp.autocast('cuda'):
                out = model(imgs)
                loss = criterion(out, lbls)
            scaler.scale(loss).backward()
            scaler.step(optimizer)
            scaler.update()
            
            t_loss += loss.item() * imgs.size(0)
            _, p = torch.max(out, 1)
            t_total += lbls.size(0)
            t_correct += (p == lbls).sum().item()

        # validation phase
        model.eval()
        v_loss, v_correct, v_total = 0.0, 0, 0
        with torch.no_grad():
            for imgs, lbls in val_loader:
                imgs, lbls = imgs.to(device), lbls.to(device)
                with torch.amp.autocast('cuda'):
                    out = model(imgs)
                    loss = criterion(out, lbls)
                v_loss += loss.item() * imgs.size(0)
                _, p = torch.max(out, 1)
                v_total += lbls.size(0)
                v_correct += (p == lbls).sum().item()

        epoch_loss = v_loss / v_total
        epoch_acc = v_correct / v_total
        
        history['train_acc'].append(t_correct / t_total)
        history['val_acc'].append(epoch_acc)
        history['train_loss'].append(t_loss / t_total)
        history['val_loss'].append(epoch_loss)
        
        scheduler.step(epoch_loss)
        print(f"Epoch {epoch+1:02d}/{epochs:02d} | Train Acc: {t_correct/t_total*100:.2f}% | Val Acc: {epoch_acc*100:.2f}% | Val Loss: {epoch_loss:.4f} | Time: {time.time()-start_t:.1f}s")

        if epoch_acc > best_acc:
            best_acc = epoch_acc
            best_model_wts = copy.deepcopy(model.state_dict())
        
        stopper(epoch_loss)
        if stopper.early_stop:
            print("Early stopping triggered.")
            break
            
    # Load best weights
    model.load_state_dict(best_model_wts)
    checkpoint_name = f"best_{name.lower().replace(' ', '_').replace('(', '').replace(')', '')}.pth"
    model_save_path = os.path.join(models_dir, checkpoint_name)
    torch.save(model.state_dict(), model_save_path)
    print(f"Saved best model weights to: {model_save_path}")
    
    # ------------------- PROFILING & METRICS -------------------
    # 1. Parameter Count (Millions)
    num_params = sum(p.numel() for p in model.parameters()) / 1e6
    
    # 2. Storage Size (MB)
    storage_size = os.path.getsize(model_save_path) / (1024 * 1024)
    
    # 3. Validation Evaluations
    all_preds, all_lbls = [], []
    model.eval()
    with torch.no_grad():
        for imgs, lbls in val_loader:
            imgs = imgs.to(device)
            out = model(imgs)
            _, p = torch.max(out, 1)
            all_preds.extend(p.cpu().numpy())
            all_lbls.extend(lbls.numpy())
            
    val_acc = accuracy_score(all_lbls, all_preds)
    macro_precision, macro_recall, macro_f1, _ = precision_recall_fscore_support(
        all_lbls, all_preds, average='macro', zero_division=0
    )
    
    # 4. Latency Profiling (GPU and CPU)
    # GPU Latency
    dummy_input_cpu = torch.randn(1, 3, 224, 224)
    dummy_input_gpu = dummy_input_cpu.to(device)
    
    # Warm-up passes
    with torch.no_grad():
        for _ in range(20):
            _ = model(dummy_input_gpu)
            
    # Measure GPU
    if torch.cuda.is_available():
        torch.cuda.synchronize()
    start_gpu = time.time()
    with torch.no_grad():
        for _ in range(100):
            _ = model(dummy_input_gpu)
    if torch.cuda.is_available():
        torch.cuda.synchronize()
    gpu_latency = ((time.time() - start_gpu) / 100.0) * 1000.0  # ms
    
    # Measure CPU
    model_cpu = copy.deepcopy(model).cpu()
    model_cpu.eval()
    # Warm-up CPU
    with torch.no_grad():
        for _ in range(10):
            _ = model_cpu(dummy_input_cpu)
    start_cpu = time.time()
    with torch.no_grad():
        for _ in range(100):
            _ = model_cpu(dummy_input_cpu)
    cpu_latency = ((time.time() - start_cpu) / 100.0) * 1000.0  # ms
    
    # 5. Peak GPU Memory (MB)
    if torch.cuda.is_available():
        torch.cuda.reset_peak_memory_stats(device)
        with torch.no_grad():
            for imgs, lbls in val_loader:
                imgs = imgs.to(device)
                _ = model(imgs)
        peak_gpu_mem = torch.cuda.max_memory_allocated(device) / (1024 * 1024)
    else:
        peak_gpu_mem = 0.0
        
    print(f"Model {name} Evaluation Results:")
    print(f"  Accuracy:  {val_acc*100:.2f}%")
    print(f"  Precision: {macro_precision*100:.2f}%")
    print(f"  Recall:    {macro_recall*100:.2f}%")
    print(f"  F1-Score:  {macro_f1*100:.2f}%")
    print(f"  Params:    {num_params:.2f} M")
    print(f"  Size:      {storage_size:.2f} MB")
    print(f"  GPU Latency: {gpu_latency:.2f} ms/frame")
    print(f"  CPU Latency: {cpu_latency:.2f} ms/frame")
    print(f"  Peak GPU Mem: {peak_gpu_mem:.2f} MB")
    
    results.append({
        "Model": name,
        "Accuracy (%)": round(val_acc * 100, 2),
        "Precision (%)": round(macro_precision * 100, 2),
        "Recall (%)": round(macro_recall * 100, 2),
        "F1-Score (%)": round(macro_f1 * 100, 2),
        "Parameters (M)": round(num_params, 2),
        "Storage Size (MB)": round(storage_size, 2),
        "GPU Latency (ms)": round(gpu_latency, 2),
        "CPU Latency (ms)": round(cpu_latency, 2),
        "Peak GPU Mem (MB)": round(peak_gpu_mem, 2)
    })
    
    # Save training curves plot
    sns.set_theme(style="whitegrid")
    plt.rcParams['figure.dpi'] = 300
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    ax1.plot(history['train_acc'], label='Train Acc', marker='o')
    ax1.plot(history['val_acc'], label='Val Acc', marker='s')
    ax1.set_title(f'{name} Accuracy Curves')
    ax1.set_xlabel('Epoch')
    ax1.set_ylabel('Accuracy')
    ax1.legend()
    
    ax2.plot(history['train_loss'], label='Train Loss', marker='o')
    ax2.plot(history['val_loss'], label='Val Loss', marker='s')
    ax2.set_title(f'{name} Loss Curves')
    ax2.set_xlabel('Epoch')
    ax2.set_ylabel('Loss')
    ax2.legend()
    
    curves_path = os.path.join(output_dir, f"{name.lower().replace(' ', '_').replace('(', '').replace(')', '')}_curves.png")
    plt.savefig(curves_path, bbox_inches='tight')
    plt.close()
    
    # Save Confusion Matrix
    cm = confusion_matrix(all_lbls, all_preds)
    plt.figure(figsize=(10, 8))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', xticklabels=classes, yticklabels=classes)
    plt.title(f'Confusion Matrix: {name}')
    plt.xlabel('Predicted')
    plt.ylabel('True')
    cm_path = os.path.join(output_dir, f"{name.lower().replace(' ', '_').replace('(', '').replace(')', '')}_cm.png")
    plt.savefig(cm_path, bbox_inches='tight')
    plt.close()
    
    # Clean up GPU memory
    del model, model_cpu
    gc.collect()
    torch.cuda.empty_cache()

# Save all results to CSV
df = pd.DataFrame(results)
csv_path = os.path.join(output_dir, "benchmark_results.csv")
df.to_csv(csv_path, index=False)
print(f"\nSaved comparison results table to: {csv_path}")

# Generate a styled comparison table image
def save_table_image(data_frame, filepath):
    fig, ax = plt.subplots(figsize=(14, len(data_frame) * 0.6 + 2.0))
    ax.axis('off')
    
    # Create the table
    tbl = ax.table(cellText=data_frame.values, colLabels=data_frame.columns, cellLoc='center', loc='center')
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1.0, 1.6)
    
    # Style header and rows
    for key, cell in tbl.get_celld().items():
        cell.set_linewidth(0.8)
        cell.set_edgecolor('#d3d3d3')
        if key[0] == 0:  # Header
            cell.set_text_props(weight='bold', color='white')
            cell.set_facecolor('#2c3e50') # Sleek dark gray blue
        else:            # Rows
            if key[0] % 2 == 0:
                cell.set_facecolor('#f9f9f9') # Light row color
            else:
                cell.set_facecolor('#ffffff') # White row color
                
            # Highlight best model (let's check which is the best accuracy model)
            best_idx = data_frame['Accuracy (%)'].idxmax() + 1
            if key[0] == best_idx:
                cell.set_text_props(weight='bold')
                # Light green accent for the best model row
                if key[1] == 0:
                    cell.set_facecolor('#d4edda')
                else:
                    cell.set_facecolor('#e8f5e9')
                    
    plt.title("Lightweight Deep Transfer Learning Models Performance Comparison", weight='bold', fontsize=12, pad=15)
    plt.savefig(filepath, dpi=300, bbox_inches='tight')
    plt.close()

table_img_path = os.path.join(output_dir, "benchmark_table.png")
save_table_image(df, table_img_path)
print(f"Saved styled comparison table image to: {table_img_path}")

# Generate trade-off plot: Accuracy vs. GPU Latency
plt.figure(figsize=(10, 6))
p1 = sns.scatterplot(
    data=df, 
    x="GPU Latency (ms)", 
    y="Accuracy (%)", 
    size="Parameters (M)", 
    hue="Model",
    sizes=(80, 400),
    palette="viridis",
    legend="brief"
)
# Add annotations
for idx, row in df.iterrows():
    plt.text(
        row["GPU Latency (ms)"] + 0.1, 
        row["Accuracy (%)"] + 0.1, 
        row["Model"], 
        horizontalalignment='left', 
        size='small', 
        color='black', 
        weight='semibold'
    )
plt.title("Trade-off Analysis: Accuracy vs. GPU Inference Latency", fontsize=12, weight='bold')
plt.xlabel("GPU Inference Latency (ms per frame)", fontsize=10)
plt.ylabel("Validation Accuracy (%)", fontsize=10)
plt.xlim(df["GPU Latency (ms)"].min() - 0.5, df["GPU Latency (ms)"].max() + 1.0)
plt.ylim(df["Accuracy (%)"].min() - 1.0, df["Accuracy (%)"].max() + 1.0)
tradeoff_path = os.path.join(output_dir, "accuracy_vs_latency_tradeoff.png")
plt.savefig(tradeoff_path, bbox_inches='tight')
plt.close()
print(f"Saved trade-off scatterplot to: {tradeoff_path}")

print("\nBenchmarking pipeline completed successfully!")
