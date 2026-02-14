"""
SRT File Parser Module
Handles reading and parsing of SRT subtitle files into structured data.
"""

import re
import os


class SubtitleEntry:
    """Represents a single subtitle entry."""

    def __init__(self, index, start_time, end_time, text):
        self.index = index
        self.start_time = start_time
        self.end_time = end_time
        self.text = text

    def __repr__(self):
        return (
            f"SubtitleEntry(index={self.index}, "
            f"start='{self.start_time}', "
            f"end='{self.end_time}', "
            f"text='{self.text[:30]}...')"
        )

    def to_dict(self):
        """Convert entry to dictionary."""
        return {
            'index': self.index,
            'start_time': self.start_time,
            'end_time': self.end_time,
            'text': self.text
        }

    def get_duration_seconds(self):
        """Calculate duration of subtitle in seconds."""
        start = self._time_to_seconds(self.start_time)
        end = self._time_to_seconds(self.end_time)
        return round(end - start, 3)

    @staticmethod
    def _time_to_seconds(time_str):
        """Convert timestamp string to seconds."""
        time_str = time_str.replace(',', '.')
        parts = time_str.split(':')
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = float(parts[2])
        return hours * 3600 + minutes * 60 + seconds


class SRTParser:
    """
    Parses SRT subtitle files into structured SubtitleEntry objects.

    Supports multiple encodings and handles malformed SRT files gracefully.
    """

    # Regex pattern for matching SRT blocks
    BLOCK_PATTERN = re.compile(
        r'(\d+)\s*\n'
        r'(\d{2}:\d{2}:\d{2}[,\.]\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}[,\.]\d{3})\s*\n'
        r'((?:(?!\d+\s*\n\d{2}:\d{2}:\d{2}).+\n?)+)',
        re.MULTILINE
    )

    # Timestamp line pattern
    TIMESTAMP_PATTERN = re.compile(
        r'(\d{2}:\d{2}:\d{2}[,\.]\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}[,\.]\d{3})'
    )

    # Supported encodings to try
    ENCODINGS = [
        'utf-8-sig',
        'utf-8',
        'latin-1',
        'cp1252',
        'iso-8859-1',
        'ascii',
        'utf-16',
        'utf-16-le',
        'utf-16-be'
    ]

    # HTML/ASS tag pattern for cleanup
    TAG_PATTERN = re.compile(r'<[^>]+>|{[^}]+}')

    def __init__(self):
        self.errors = []

    def parse_file(self, file_path):
        """
        Parse an SRT file and return a list of SubtitleEntry objects.

        Args:
            file_path: Path to the SRT file

        Returns:
            List of SubtitleEntry objects

        Raises:
            FileNotFoundError: If file does not exist
            ValueError: If file cannot be read with any encoding
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        if not file_path.lower().endswith('.srt'):
            raise ValueError(f"Not an SRT file: {file_path}")

        self.errors = []
        content = self._read_file(file_path)
        content = self._clean_content(content)

        # Try regex-based parsing first
        subtitles = self._regex_parse(content)

        # Fallback to block-based parsing if regex fails
        if not subtitles:
            subtitles = self._block_parse(content)

        return subtitles

    def _read_file(self, file_path):
        """
        Read file content trying multiple encodings.

        Args:
            file_path: Path to the file

        Returns:
            File content as string

        Raises:
            ValueError: If no encoding works
        """
        for encoding in self.ENCODINGS:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                return content
            except (UnicodeDecodeError, UnicodeError):
                continue

        raise ValueError(
            f"Could not read file '{file_path}' with any known encoding. "
            f"Tried: {', '.join(self.ENCODINGS)}"
        )

    def _clean_content(self, content):
        """
        Clean and normalize file content.

        Args:
            content: Raw file content

        Returns:
            Cleaned content string
        """
        # Normalize line endings
        content = content.replace('\r\n', '\n')
        content = content.replace('\r', '\n')

        # Remove BOM if present
        content = content.lstrip('\ufeff')

        # Strip leading/trailing whitespace
        content = content.strip()

        return content

    def _clean_subtitle_text(self, text):
        """
        Clean subtitle text by removing HTML/ASS tags.

        Args:
            text: Raw subtitle text

        Returns:
            Cleaned text string
        """
        # Remove HTML and ASS style tags
        text = self.TAG_PATTERN.sub('', text)

        # Clean up extra whitespace
        text = text.strip()

        # Remove empty lines within text
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        text = '\n'.join(lines)

        return text

    def _regex_parse(self, content):
        """
        Parse content using regex pattern matching.

        Args:
            content: Cleaned file content

        Returns:
            List of SubtitleEntry objects
        """
        subtitles = []
        matches = self.BLOCK_PATTERN.findall(content)

        for match in matches:
            try:
                index = int(match[0])
                start_time = match[1].replace(',', '.')
                end_time = match[2].replace(',', '.')
                text = self._clean_subtitle_text(match[3])

                if text:
                    entry = SubtitleEntry(index, start_time, end_time, text)
                    subtitles.append(entry)
            except (ValueError, IndexError) as e:
                self.errors.append(f"Error parsing block {match[0]}: {str(e)}")
                continue

        return subtitles

    def _block_parse(self, content):
        """
        Fallback block-based parsing for malformed SRT files.

        Args:
            content: Cleaned file content

        Returns:
            List of SubtitleEntry objects
        """
        subtitles = []
        blocks = content.split('\n\n')

        for block in blocks:
            block = block.strip()
            if not block:
                continue

            lines = block.split('\n')

            if len(lines) < 3:
                continue

            # Parse index number
            try:
                index = int(lines[0].strip())
            except ValueError:
                continue

            # Parse timestamp line
            time_match = self.TIMESTAMP_PATTERN.match(lines[1].strip())
            if not time_match:
                continue

            start_time = time_match.group(1).replace(',', '.')
            end_time = time_match.group(2).replace(',', '.')

            # Parse text (everything after timestamp)
            text_lines = lines[2:]
            text = '\n'.join(text_lines)
            text = self._clean_subtitle_text(text)

            if text:
                entry = SubtitleEntry(index, start_time, end_time, text)
                subtitles.append(entry)

        return subtitles

    def get_file_info(self, file_path):
        """
        Get basic info about an SRT file without full parsing.

        Args:
            file_path: Path to the SRT file

        Returns:
            Dictionary with file information
        """
        stat = os.stat(file_path)

        return {
            'filename': os.path.basename(file_path),
            'path': file_path,
            'size_bytes': stat.st_size,
            'size_formatted': self._format_size(stat.st_size),
        }

    @staticmethod
    def _format_size(size_bytes):
        """Format byte count to human-readable string."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"

    @staticmethod
    def find_srt_files(folder_path, recursive=True):
        """
        Find all SRT files in a folder.

        Args:
            folder_path: Path to search
            recursive: Whether to search subdirectories

        Returns:
            Sorted list of SRT file paths
        """
        srt_files = []

        if recursive:
            for root_dir, dirs, files in os.walk(folder_path):
                for filename in files:
                    if filename.lower().endswith('.srt'):
                        full_path = os.path.join(root_dir, filename)
                        srt_files.append(full_path)
        else:
            for filename in os.listdir(folder_path):
                if filename.lower().endswith('.srt'):
                    full_path = os.path.join(folder_path, filename)
                    if os.path.isfile(full_path):
                        srt_files.append(full_path)

        return sorted(srt_files)