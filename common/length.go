// Package common provides shared value types used across all OOXML format
// packages (docx, pptx, xlsx).
package common

const emuPerInch = 914400
const emuPerCm = 360000
const emuPerPt = 12700
const emuPerTwip = 635

// Length represents a distance in English Metric Units (EMU).
// There are 914,400 EMU per inch. Length is used for page dimensions,
// margins, column widths, image sizes, and other measurements in OOXML.
//
// Create a Length using one of the constructor functions: [EMU], [Inches],
// [Cm], [Pt], or [Twips]. Convert back using the corresponding methods.
//
// Length is a value type and is safe for concurrent use.
type Length int64

// EMU creates a [Length] from English Metric Units.
func EMU(v int64) Length { return Length(v) }

// Inches creates a [Length] from inches. One inch equals 914,400 EMU.
func Inches(v float64) Length { return Length(int64(v * emuPerInch)) }

// Cm creates a [Length] from centimeters. One centimeter equals 360,000 EMU.
func Cm(v float64) Length { return Length(int64(v * emuPerCm)) }

// Pt creates a [Length] from typographic points. One point equals 1/72 inch (12,700 EMU).
func Pt(v float64) Length { return Length(int64(v * emuPerPt)) }

// Twips creates a [Length] from twips. One twip equals 1/1440 inch (635 EMU).
func Twips(v int64) Length { return Length(v * emuPerTwip) }

// EMU returns the length in English Metric Units.
func (l Length) EMU() int64 { return int64(l) }

// Inches returns the length in inches.
func (l Length) Inches() float64 { return float64(l) / emuPerInch }

// Cm returns the length in centimeters.
func (l Length) Cm() float64 { return float64(l) / emuPerCm }

// Pt returns the length in typographic points.
func (l Length) Pt() float64 { return float64(l) / emuPerPt }

// Twips returns the length in twips (1/1440 inch).
func (l Length) Twips() int64 { return int64(l) / emuPerTwip }
