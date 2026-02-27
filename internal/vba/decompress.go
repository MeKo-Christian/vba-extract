package vba

import (
	"encoding/binary"
	"fmt"
)

const (
	compressedContainerSig = 0x01
	maxChunkDecompressed   = 4096
)

type DecompressionStrategy string

const (
	StrategyStandard    DecompressionStrategy = "standard"
	StrategySkipPrefix  DecompressionStrategy = "skip-prefix"
	StrategyRawPassthru DecompressionStrategy = "raw-passthrough"
)

var bitCountLUT = buildBitCountLUT()

// DecompressContainer decompresses a VBA compressed container as specified by
// MS-OVBA §2.4.1.
func DecompressContainer(data []byte) ([]byte, error) {
	if len(data) == 0 {
		return nil, fmt.Errorf("vba: empty compressed container")
	}
	if data[0] != compressedContainerSig {
		return nil, fmt.Errorf("vba: invalid compressed container signature %#x", data[0])
	}

	var out []byte
	pos := 1

	for pos < len(data) {
		if pos+2 > len(data) {
			return nil, fmt.Errorf("vba: truncated chunk header at offset %d", pos)
		}

		header := binary.LittleEndian.Uint16(data[pos : pos+2])
		pos += 2

		sigBits := (header >> 12) & 0x7
		if sigBits != 0x3 {
			return nil, fmt.Errorf("vba: invalid chunk signature bits %#x", sigBits)
		}

		compressed := (header & 0x8000) != 0
		chunkSize := int(header&0x0FFF) + 3 // includes 2-byte header
		if chunkSize < 3 {
			return nil, fmt.Errorf("vba: invalid chunk size %d", chunkSize)
		}

		payloadLen := chunkSize - 2
		if pos+payloadLen > len(data) {
			return nil, fmt.Errorf("vba: truncated chunk payload at offset %d (need %d bytes)", pos, payloadLen)
		}

		payload := data[pos : pos+payloadLen]
		pos += payloadLen

		if !compressed {
			out = append(out, payload...)
			continue
		}

		chunkOut, err := decompressChunk(payload)
		if err != nil {
			return nil, err
		}
		out = append(out, chunkOut...)
	}

	return out, nil
}

// DecompressContainerWithFallback tries multiple strategies in this order:
// 1) Standard decompression from byte 0
// 2) Skip leading bytes and retry from first compressed signature byte (0x01)
// 3) Return raw bytes as-is
func DecompressContainerWithFallback(data []byte, verbose bool, logf func(string, ...interface{})) ([]byte, DecompressionStrategy, error) {
	if len(data) == 0 {
		return nil, "", fmt.Errorf("vba: empty input for decompression fallback")
	}

	out, err := DecompressContainer(data)
	if err == nil {
		if verbose && logf != nil {
			logf("vba: decompression strategy=%s", StrategyStandard)
		}
		return out, StrategyStandard, nil
	}

	if verbose && logf != nil {
		logf("vba: standard decompression failed: %v", err)
	}

	for i := 1; i < len(data); i++ {
		if data[i] != compressedContainerSig {
			continue
		}

		out, skipErr := DecompressContainer(data[i:])
		if skipErr != nil {
			continue
		}

		if verbose && logf != nil {
			logf("vba: decompression strategy=%s offset=%d", StrategySkipPrefix, i)
		}
		return out, StrategySkipPrefix, nil
	}

	raw := make([]byte, len(data))
	copy(raw, data)
	if verbose && logf != nil {
		logf("vba: decompression strategy=%s", StrategyRawPassthru)
	}
	return raw, StrategyRawPassthru, nil
}

func decompressChunk(payload []byte) ([]byte, error) {
	out := make([]byte, 0, maxChunkDecompressed)
	pos := 0

	for pos < len(payload) {
		flags := payload[pos]
		pos++

		for bit := 0; bit < 8; bit++ {
			if pos >= len(payload) {
				break
			}

			isCopy := ((flags >> bit) & 0x01) == 0x01
			if !isCopy {
				out = append(out, payload[pos])
				pos++
				if len(out) > maxChunkDecompressed {
					return nil, fmt.Errorf("vba: decompressed chunk exceeds %d bytes", maxChunkDecompressed)
				}
				continue
			}

			if pos+2 > len(payload) {
				return nil, fmt.Errorf("vba: truncated copy token in compressed chunk")
			}

			token := binary.LittleEndian.Uint16(payload[pos : pos+2])
			pos += 2

			bitCount := bitCountForDecompressedPos(len(out))
			lengthBits := 16 - bitCount
			lengthMask := uint16((1 << lengthBits) - 1)

			offset := int(token>>lengthBits) + 1
			length := int(token&lengthMask) + 3

			if offset <= 0 || offset > len(out) {
				return nil, fmt.Errorf("vba: invalid copy offset %d for decompressed length %d", offset, len(out))
			}

			for i := 0; i < length; i++ {
				src := len(out) - offset
				out = append(out, out[src])
				if len(out) > maxChunkDecompressed {
					return nil, fmt.Errorf("vba: decompressed chunk exceeds %d bytes", maxChunkDecompressed)
				}
			}
		}
	}

	return out, nil
}

func bitCountForDecompressedPos(pos int) int {
	if pos < 0 {
		return 4
	}
	if pos >= len(bitCountLUT) {
		return 12
	}
	return bitCountLUT[pos]
}

func buildBitCountLUT() []int {
	lut := make([]int, maxChunkDecompressed+1)
	for i := 0; i <= maxChunkDecompressed; i++ {
		switch {
		case i <= 16:
			lut[i] = 4
		case i <= 32:
			lut[i] = 5
		case i <= 64:
			lut[i] = 6
		case i <= 128:
			lut[i] = 7
		case i <= 256:
			lut[i] = 8
		case i <= 512:
			lut[i] = 9
		case i <= 1024:
			lut[i] = 10
		case i <= 2048:
			lut[i] = 11
		default:
			lut[i] = 12
		}
	}
	return lut
}
