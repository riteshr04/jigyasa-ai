from src.train_model import train_and_evaluate


def main():
    model, metrics = train_and_evaluate()
    print("Training completed successfully.")
    print(f"Model type: {type(model).__name__}")
    print(f"Accuracy: {metrics['accuracy']:.4f}")
    print(metrics["classification_report"])


if __name__ == "__main__":
    main()
