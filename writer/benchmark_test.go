package writer

import (
	"bytes"
	"strings"
	"testing"
)

// BenchmarkBufferVsBuilder compares bytes.Buffer vs strings.Builder performance
func BenchmarkBufferVsBuilder(b *testing.B) {
	testData := []string{"Hello", " ", "World", "!", " ", "This", " ", "is", " ", "a", " ", "test"}

	b.Run("BytesBuffer", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var buf bytes.Buffer
			for _, s := range testData {
				buf.WriteString(s)
			}
			_ = buf.String()
		}
	})

	b.Run("StringsBuilder", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var builder strings.Builder
			for _, s := range testData {
				builder.WriteString(s)
			}
			_ = builder.String()
		}
	})

	b.Run("StringsBuilderPrealloc", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var builder strings.Builder
			builder.Grow(64) // Pre-allocate capacity
			for _, s := range testData {
				builder.WriteString(s)
			}
			_ = builder.String()
		}
	})
}

// BenchmarkBufferPool tests the performance of buffer pooling
func BenchmarkBufferPool(b *testing.B) {
	testData := []byte("This is test data for buffer pool benchmarking")

	b.Run("NewBuffer", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var buf bytes.Buffer
			buf.Write(testData)
			_ = buf.Bytes()
		}
	})

	b.Run("BufferPool", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			buf := getBuffer()
			buf.Write(testData)
			result := make([]byte, buf.Len())
			copy(result, buf.Bytes())
			putBuffer(buf)
			_ = result
		}
	})
}

// BenchmarkXMLGeneration compares different XML generation approaches
func BenchmarkXMLGeneration(b *testing.B) {
	b.Run("StringConcatenation", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			xml := "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
			xml += "<root>\n"
			xml += "  <element>value</element>\n"
			xml += "  <element>value2</element>\n"
			xml += "</root>"
			_ = xml
		}
	})

	b.Run("StringsBuilder", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var builder strings.Builder
			builder.Grow(128)
			builder.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
			builder.WriteString("<root>\n")
			builder.WriteString("  <element>value</element>\n")
			builder.WriteString("  <element>value2</element>\n")
			builder.WriteString("</root>")
			_ = builder.String()
		}
	})

	b.Run("StreamingXMLWriter", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			var buf bytes.Buffer
			writer := NewStreamingXMLWriter(&buf)
			writer.WriteHeader()
			writer.WriteStartElement("root")
			writer.WriteElement("element", "value")
			writer.WriteElement("element", "value2")
			writer.WriteEndElement("root")
			writer.Close()
			_ = buf.String()
		}
	})
}

// BenchmarkMemoryAllocation measures memory allocation patterns
func BenchmarkMemoryAllocation(b *testing.B) {
	b.Run("MultipleSmallAllocs", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			var results [][]byte
			for j := 0; j < 10; j++ {
				data := make([]byte, 100)
				results = append(results, data)
			}
			_ = results
		}
	})

	b.Run("SingleLargeAlloc", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			data := make([]byte, 1000)
			var results [][]byte
			for j := 0; j < 10; j++ {
				start := j * 100
				end := start + 100
				results = append(results, data[start:end])
			}
			_ = results
		}
	})

	b.Run("BufferPoolAlloc", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			buf := getBuffer()
			for j := 0; j < 10; j++ {
				buf.Write([]byte("test data"))
			}
			result := make([]byte, buf.Len())
			copy(result, buf.Bytes())
			putBuffer(buf)
			_ = result
		}
	})
}