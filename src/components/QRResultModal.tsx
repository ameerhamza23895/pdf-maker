import { electricCuratorTheme } from "@/src/theme/electric-curator";
import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import * as Clipboard from "expo-clipboard";
import React from "react";
import { Modal, StyleSheet, Text, TouchableOpacity, View } from "react-native";

const { colors, spacing, radius, typography } = electricCuratorTheme;

interface QRResultModalProps {
  isVisible: boolean;
  data: string | null;
  onClose: () => void;
}

export function QRResultModal({
  isVisible,
  data,
  onClose,
}: QRResultModalProps) {
  const copyToClipboard = async () => {
    if (data) {
      await Clipboard.setStringAsync(data);
      onClose(); // Auto-close after copying
    }
  };

  return (
    <Modal
      animationType="fade"
      transparent={true}
      visible={isVisible}
      onRequestClose={onClose}
    >
      <View style={styles.overlay}>
        <View style={styles.content}>
          <TouchableOpacity style={styles.closeButton} onPress={onClose}>
            <MaterialIcons name="close" size={20} color={colors.onSurface} />
          </TouchableOpacity>

          <View style={styles.iconCircle}>
            <MaterialIcons
              name="qr-code-scanner"
              size={28}
              color={colors.primary}
            />
          </View>

          <Text style={styles.title}>QR Code Detected</Text>

          <View style={styles.dataContainer}>
            <Text style={styles.dataText} numberOfLines={4}>
              {data}
            </Text>
          </View>

          <TouchableOpacity style={styles.copyButton} onPress={copyToClipboard}>
            <MaterialIcons
              name="content-copy"
              size={18}
              color={colors.onPrimary}
            />
            <Text style={styles.copyButtonText}>Copy to Clipboard</Text>
          </TouchableOpacity>

          <TouchableOpacity style={styles.cancelButton} onPress={onClose}>
            <Text style={styles.cancelText}>Dismiss</Text>
          </TouchableOpacity>
        </View>
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  overlay: {
    flex: 1,
    backgroundColor: "rgba(0, 0, 0, 0.6)",
    justifyContent: "center",
    alignItems: "center",
    padding: spacing.lg,
  },
  content: {
    width: "100%",
    maxWidth: 320,
    backgroundColor: colors.surface,
    borderRadius: radius.md,
    padding: spacing.lg,
    alignItems: "center",
    elevation: 20,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 10 },
    shadowOpacity: 0.3,
    shadowRadius: 20,
  },
  iconCircle: {
    width: 60,
    height: 60,
    borderRadius: 30,
    backgroundColor: colors.primaryContainer,
    justifyContent: "center",
    alignItems: "center",
    marginBottom: spacing.md,
  },
  title: {
    ...typography.labelMd,
    fontSize: 20,
    fontWeight: "700",
    color: colors.onSurface,
    marginBottom: spacing.sm,
  },
  dataContainer: {
    width: "100%",
    backgroundColor: colors.surfaceContainerLow,
    padding: spacing.md,
    borderRadius: radius.md,
    marginBottom: spacing.lg,
  },
  dataText: {
    textAlign: "center",
    color: "#000",
    fontSize: 20,
    lineHeight: 20,
  },
  copyButton: {
    flexDirection: "row",
    backgroundColor: colors.primary,
    paddingVertical: 14,
    paddingHorizontal: 24,
    borderRadius: radius.pill,
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    gap: 10,
  },
  copyButtonText: {
    color: colors.onPrimary,
    fontWeight: "600",
    fontSize: 15,
  },
  cancelButton: {
    marginTop: spacing.md,
    padding: spacing.xs,
  },
  closeButton: {
    position: "absolute",
    top: spacing.md,
    right: spacing.md,
    width: 40,
    height: 40,
    borderRadius: radius.pill,
    backgroundColor: colors.surfaceContainerLowest,
    alignItems: "center",
    justifyContent: "center",
    zIndex: 10,
  },
  cancelText: {
    color: colors.primary,
    fontWeight: "500",
  },
});
